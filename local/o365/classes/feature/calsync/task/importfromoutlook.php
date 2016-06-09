<?php
// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * @package local_o365
 * @author James McQuillan <james.mcquillan@remote-learner.net>
 * @license http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 * @copyright (C) 2014 onwards Microsoft, Inc. (http://microsoft.com/)
 */

namespace local_o365\feature\calsync\task;

/**
 * Scheduled task to check for new o365 events and sync them into Moodle.
 */
class importfromoutlook extends \core\task\scheduled_task {
    /**
     * Get a descriptive name for this task (shown to admins).
     *
     * @return string
     */
    public function get_name() {
        return get_string('task_calendarsyncin', 'local_o365');
    }

    /**
     * Do the job.
     */
    public function execute() {
        global $DB, $CFG;
        require_once($CFG->dirroot.'/calendar/lib.php');

        if (\local_o365\utils::is_configured() !== true) {
            return false;
        }

        // Get calendars set to sync in.
        $starttime = time();

        \local_o365\feature\calsync\observers::set_event_import(true);

        // Using a direct query here so we don't run into static cache issues.
        $laststarttime = $DB->get_record('config_plugins', ['plugin' => 'local_o365', 'name' => 'calsyncinlastrun']);
        $laststarttime = (!empty($laststarttime) && !empty($laststarttime->value)) ? $laststarttime->value : 0;

        $httpclient = new \local_o365\httpclient();
        $clientdata = \local_o365\oauth2\clientdata::instance_from_oidc();

        $calsubs = $DB->get_recordset_select('local_o365_calsub', 'syncbehav = ? OR syncbehav = ?', ['in', 'both']);
        $calsync = new \local_o365\feature\calsync\main($clientdata, $httpclient);
        foreach ($calsubs as $i => $calsub) {
            try {
                mtrace('Syncing events for user #'.$calsub->user_id);
                $events = $calsync->get_events($calsub->user_id, $calsub->o365calid, $laststarttime);
                if (!empty($events) && is_array($events) && isset($events['value']) && is_array($events['value'])) {
                    if (!empty($events['value'])) {
                        foreach ($events['value'] as $i => $event) {
                            if (!isset($event['Id'])) {
                                $errmsg = 'Skipped an event because of malformed data.';
                                \local_o365\utils::debug($errmsg, 'importfromoutlook', $event);
                                mtrace($errmsg);
                                continue;
                            }
                            // Verify the Outlook calendar subscription type is an Outlook Group or standard calandar.
                            if (!empty($calsub->o365calemail) && validate_email($calsub->o365calemail)) {
                                $idmapexists = false;
                                if ($calsub->o365calemail == $event['organizer']['emailAddress']['address']) {
                                    // Perform sync from Outlook group event to Moodle event.
                                    $idmapexists = $DB->record_exists('local_o365_calidmap', ['outlookeventid' => $event['Id']]);
                                } else {
                                    // If the event organizer email does not equal the subscriptions email, then ignore this event.
                                    $idmapexists = true;
                                }
                            } else {
                                $idmapexists = $DB->record_exists('local_o365_calidmap', ['outlookeventid' => $event['Id']]);
                            }
                            if ($idmapexists === false) {
                                // Create Moodle event.
                                $eventparams = [
                                    'name' => $event['Subject'],
                                    'description' => $event['Body']['Content'],
                                    'eventtype' => $calsub->caltype,
                                    'repeatid' => 0,
                                    'modulename' => 0,
                                    'instance' => 0,
                                    'timestart' => strtotime($event['Start']),
                                    'visible' => 1,
                                    'uuid' => '',
                                    'sequence' => 1,
                                ];
                                $end = strtotime($event['End']);
                                $eventparams['timeduration'] = $end - $eventparams['timestart'];

                                if ($calsub->caltype === 'user') {
                                    $eventparams['userid'] = $calsub->caltypeid;
                                }
                                if ($calsub->caltype === 'course') {
                                    $eventparams['courseid'] = $calsub->caltypeid;
                                }
                                $moodleevent = \calendar_event::create($eventparams, false);
                                if (!empty($moodleevent) && !empty($moodleevent->id)) {
                                    $idmaprec = [
                                        'eventid' => $moodleevent->id,
                                        'outlookeventid' => $event['Id'],
                                        'origin' => 'o365',
                                        'userid' => $calsub->user_id
                                    ];
                                    $DB->insert_record('local_o365_calidmap', (object)$idmaprec);
                                    mtrace('Successfully imported event #'.$moodleevent->id);
                                }
                            }
                        }
                    } else {
                        mtrace('No new events to sync in.');
                    }
                } else {
                    $errmsg = 'Bad response received when fetching events.';
                    \local_o365\utils::debug($errmsg, 'importfromoutlook', $events);
                    mtrace($errmsg);
                }
            } catch (\Exception $e) {
                \local_o365\utils::debug('Error syncing events: '.$e->getMessage(), 'importfromoutlook', $e);
                mtrace('Error: '.$e->getMessage());
            }
        }
        $calsubs->close();
        \local_o365\feature\calsync\observers::set_event_import(false);

        set_config('calsyncinlastrun', $starttime, 'local_o365');
        return true;
    }
}
