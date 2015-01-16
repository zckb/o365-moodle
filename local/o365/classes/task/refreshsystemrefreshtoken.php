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
 * @copyright (C) 2014 onwards Microsoft Open Technologies, Inc. (http://msopentech.com/)
 */

namespace local_o365\task;

/**
 * Scheduled task to refresh the system API user's refresh token.
 */
class refreshsystemrefreshtoken extends \core\task\scheduled_task {
    /**
     * Get a descriptive name for this task (shown to admins).
     *
     * @return string
     */
    public function get_name() {
        return get_string('task_refreshsystemrefreshtoken', 'local_o365');
    }

    /**
     * Do the job.
     */
    public function execute() {
        // Attempt token refresh.
        $oidcconfig = get_config('auth_oidc');
        if (!empty($oidcconfig)) {
            $httpclient = new \local_o365\httpclient();
            $clientdata = new \local_o365\oauth2\clientdata($oidcconfig->clientid, $oidcconfig->clientsecret,
                    $oidcconfig->authendpoint, $oidcconfig->tokenendpoint);
            $graphresource = 'https://graph.windows.net';
            $systemtoken = \local_o365\oauth2\systemtoken::get_for_new_resource($graphresource, $clientdata, $httpclient);
        }
        return true;
    }
}