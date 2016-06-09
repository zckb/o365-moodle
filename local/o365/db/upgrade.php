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

/**
 * Update plugin.
 *
 * @param int $oldversion the version we are upgrading from
 * @return bool result
 */
function xmldb_local_o365_upgrade($oldversion) {
    global $DB;

    $dbman = $DB->get_manager();
    $result = true;

    if ($result && $oldversion < 2014111700) {
        if (!$dbman->table_exists('local_o365_token')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_token');
        }
        upgrade_plugin_savepoint($result, '2014111700', 'local', 'o365');
    }

    if ($result && $oldversion < 2014111702) {
        if (!$dbman->table_exists('local_o365_calsub')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_calsub');
        }
        if (!$dbman->table_exists('local_o365_calidmap')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_calidmap');
        }
        upgrade_plugin_savepoint($result, '2014111702', 'local', 'o365');
    }

    if ($result && $oldversion < 2014111703) {
        $table = new xmldb_table('local_o365_calidmap');
        $field = new xmldb_field('outlookeventid', XMLDB_TYPE_CHAR, '255', null, XMLDB_NOTNULL, null, null, 'eventid');
        $dbman->change_field_type($table, $field);
        upgrade_plugin_savepoint($result, '2014111703', 'local', 'o365');
    }

    if ($result && $oldversion < 2014111707) {
        if (!$dbman->table_exists('local_o365_cronqueue')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_cronqueue');
        }
        upgrade_plugin_savepoint($result, '2014111707', 'local', 'o365');
    }

    if ($result && $oldversion < 2014111710) {
        if (!$dbman->table_exists('local_o365_coursespsite')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_coursespsite');
        }
        if (!$dbman->table_exists('local_o365_spgroupdata')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_spgroupdata');
        }
        if (!$dbman->table_exists('local_o365_aaduserdata')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_aaduserdata');
        }
        upgrade_plugin_savepoint($result, '2014111710', 'local', 'o365');
    }

    if ($result && $oldversion < 2014111711) {
        $table = new xmldb_table('local_o365_spgroupdata');
        $field = new xmldb_field('permtype', XMLDB_TYPE_CHAR, '255', null, XMLDB_NOTNULL, null, null, 'grouptitle');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }
        upgrade_plugin_savepoint($result, '2014111711', 'local', 'o365');
    }

    if ($result && $oldversion < 2014111715) {
        if (!$dbman->table_exists('local_o365_spgroupassign')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_spgroupassign');
        }
        upgrade_plugin_savepoint($result, '2014111715', 'local', 'o365');
    }

    if ($result && $oldversion < 2014111716) {
        // Drop old index.
        $table = new xmldb_table('local_o365_token');
        $index = new xmldb_index('usrresscp', XMLDB_INDEX_NOTUNIQUE, ['user_id', 'resource', 'scope']);
        if ($dbman->index_exists($table, $index)) {
            $dbman->drop_index($table, $index);
        }

        // Lengthen field.
        $table = new xmldb_table('local_o365_token');
        $field = new xmldb_field('scope', XMLDB_TYPE_CHAR, '255', null, XMLDB_NOTNULL, null, null, 'user_id');
        $dbman->change_field_type($table, $field);

        // Create new index.
        $table = new xmldb_table('local_o365_token');
        $index = new xmldb_index('usrres', XMLDB_INDEX_NOTUNIQUE, ['user_id', 'resource']);
        if (!$dbman->index_exists($table, $index)) {
            $dbman->add_index($table, $index);
        }

        upgrade_plugin_savepoint($result, '2014111716', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012702) {
        // Migrate settings.
        $config = get_config('local_o365');
        if (empty($config->sharepointlink) && isset($config->tenant) && isset($config->parentsiteuri)) {
            $sharepointlink = 'https://'.$config->tenant.'.sharepoint.com/'.$config->parentsiteuri;
            set_config('sharepointlink', $sharepointlink, 'local_o365');
        }
        upgrade_plugin_savepoint($result, '2015012702', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012704) {
        $config = get_config('local_o365');
        if (!empty($config->tenant)) {
            set_config('aadtenant', $config->tenant.'.onmicrosoft.com', 'local_o365');
            set_config('odburl', $config->tenant.'-my.sharepoint.com', 'local_o365');
        }
        upgrade_plugin_savepoint($result, '2015012704', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012707) {
        $table = new xmldb_table('local_o365_calsub');
        $field = new xmldb_field('o365calid', XMLDB_TYPE_CHAR, '255', null, XMLDB_NOTNULL, null, null, 'caltypeid');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }
        $field = new xmldb_field('syncbehav', XMLDB_TYPE_CHAR, '255', null, XMLDB_NOTNULL, null, null, 'o365calid');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }
        upgrade_plugin_savepoint($result, '2015012707', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012708) {
        $table = new xmldb_table('local_o365_calidmap');
        $field = new xmldb_field('origin', XMLDB_TYPE_CHAR, '255', null, XMLDB_NOTNULL, null, null, 'outlookeventid');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }

        $idmaps = $DB->get_recordset('local_o365_calidmap');
        foreach ($idmaps as $idmap) {
            $newidmap = new \stdClass;
            $newidmap->id = $idmap->id;
            if (empty($idmap->origin)) {
                $newidmap->origin = 'moodle';
                $DB->update_record('local_o365_calidmap', $newidmap);
            }
        }
        $idmaps->close();

        upgrade_plugin_savepoint($result, '2015012708', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012709) {
        $table = new xmldb_table('local_o365_calidmap');
        $field = new xmldb_field('userid', XMLDB_TYPE_INTEGER, '10', null, XMLDB_NOTNULL, null, '0', 'origin');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }

        $table = new xmldb_table('local_o365_calsub');
        $field = new xmldb_field('isprimary', XMLDB_TYPE_INTEGER, '1', null, XMLDB_NOTNULL, null, '1', 'o365calid');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }

        // Update the calidmap with the user id of the user who created the event. Before multi-cal syncing, the o365 event was
        // always created in the calendar of the event creating user (with others as attendees).
        $sql = 'SELECT idmap.id as idmapid,
                       ev.userid as eventuserid
                  FROM {local_o365_calidmap} idmap
                  JOIN {event} ev ON idmap.eventid = ev.id';
        $idmaps = $DB->get_recordset_sql($sql);
        foreach ($idmaps as $idmap) {
            $newidmap = new \stdClass;
            if (empty($idmap->userid)) {
                $newidmap->id = $idmap->idmapid;
                $newidmap->userid = $idmap->eventuserid;
                $DB->update_record('local_o365_calidmap', $newidmap);
            }
        }
        $idmaps->close();

        upgrade_plugin_savepoint($result, '2015012709', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012710) {
        $calsubs = $DB->get_recordset('local_o365_calsub');
        foreach ($calsubs as $i => $calsub) {
            if (empty($calsub->syncbehav)) {
                $newcalsub = new \stdClass;
                $newcalsub->id = $calsub->id;
                $newcalsub->syncbehav = 'out';
                $DB->update_record('local_o365_calsub', $newcalsub);
            }
        }
        $calsubs->close();
        upgrade_plugin_savepoint($result, '2015012710', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012712) {
        // Lengthen field.
        $table = new xmldb_table('local_o365_token');
        $field = new xmldb_field('scope', XMLDB_TYPE_TEXT, null, null, null, null, null, 'user_id');
        $dbman->change_field_type($table, $field);
        upgrade_plugin_savepoint($result, '2015012712', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012713) {
        if (!$dbman->table_exists('local_o365_objects')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_objects');
        }
        upgrade_plugin_savepoint($result, '2015012713', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012714) {
        $table = new xmldb_table('local_o365_objects');
        $field = new xmldb_field('subtype', XMLDB_TYPE_CHAR, '255', null, XMLDB_NOTNULL, null, null, 'type');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }
        upgrade_plugin_savepoint($result, '2015012714', 'local', 'o365');
    }

    if ($result && $oldversion < 2015012715) {
        // Lengthen field.
        $table = new xmldb_table('local_o365_token');
        $field = new xmldb_field('scope', XMLDB_TYPE_TEXT, null, null, null, null, null, 'user_id');
        $dbman->change_field_type($table, $field);
        upgrade_plugin_savepoint($result, '2015012715', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060102) {
        $usersync = get_config('local_o365', 'aadsync');
        if ($usersync === '1') {
            set_config('aadsync', 'create', 'local_o365');
        }
        upgrade_plugin_savepoint($result, '2015060102', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060103) {
        if (!$dbman->table_exists('local_o365_connections')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_connections');
        }
        upgrade_plugin_savepoint($result, '2015060103', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060104) {
        $table = new xmldb_table('local_o365_connections');
        $field = new xmldb_field('uselogin', XMLDB_TYPE_INTEGER, '1', null, XMLDB_NOTNULL, null, '0', 'aadupn');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }
        upgrade_plugin_savepoint($result, '2015060104', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060109) {
        if ($dbman->table_exists('local_o365_aaduserdata')) {
            $now = time();
            $aaduserdatars = $DB->get_recordset('local_o365_aaduserdata');
            foreach ($aaduserdatars as $aaduserdatarec) {
                $objectrec = (object)[
                    'type' => 'user',
                    'subtype' => '',
                    'objectid' => $aaduserdatarec->objectid,
                    'moodleid' => $aaduserdatarec->muserid,
                    'o365name' => $aaduserdatarec->userupn,
                    'timecreated' => $now,
                    'timemodified' => $now,
                ];
                $objectrec->id = $DB->insert_record('local_o365_objects', $objectrec);
                if (!empty($objectrec->id)) {
                    $DB->delete_records('local_o365_aaduserdata', ['id' => $aaduserdatarec->id]);
                }
            }
        }
        upgrade_plugin_savepoint($result, '2015060109', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060111) {
        // Clean up old "calendarsyncin" task record, if present. Replaced by \local_o365\feature\calsync\task\importfromoutlook.
        $conditions = ['component' => 'local_o365', 'classname' => '\local_o365\task\calendarsyncin'];
        $DB->delete_records('task_scheduled', $conditions);
        upgrade_plugin_savepoint($result, '2015060111', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060115.01) {
        $fieldmapconfig = $DB->get_record('config_plugins', ['plugin' => 'local_o365', 'name' => 'fieldmap']);
        if (empty($fieldmapconfig)) {
            $fieldmapdefault = [
                'givenName/firstname/always',
                'surname/lastname/always',
                'mail/email/always',
                'city/city/always',
                'country/country/always',
                'department/department/always',
                'preferredLanguage/lang/always',
            ];
            set_config('fieldmap', serialize($fieldmapdefault), 'local_o365');
        }
        upgrade_plugin_savepoint($result, '2015060115.01', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060115.02) {
        $table = new xmldb_table('local_o365_appassign');
        if (!$dbman->table_exists($table)) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_appassign');
        }
        upgrade_plugin_savepoint($result, '2015060115.02', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060116.01) {
        $table = new xmldb_table('local_o365_matchqueue');
        if (!$dbman->table_exists($table)) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_matchqueue');
        }
        upgrade_plugin_savepoint($result, '2015060116.01', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060116.03) {
        // Create new indexes.
        $table = new xmldb_table('auth_oidc_token');
        $index = new xmldb_index('oidcusername', XMLDB_INDEX_NOTUNIQUE, ['oidcusername']);
        if (!$dbman->index_exists($table, $index)) {
            $dbman->add_index($table, $index);
        }
        $table = new xmldb_table('local_o365_matchqueue');
        $index = new xmldb_index('completed', XMLDB_INDEX_NOTUNIQUE, ['completed']);
        if (!$dbman->index_exists($table, $index)) {
            $dbman->add_index($table, $index);
        }
        $index = new xmldb_index('o365username', XMLDB_INDEX_NOTUNIQUE, ['o365username']);
        if (!$dbman->index_exists($table, $index)) {
            $dbman->add_index($table, $index);
        }
        upgrade_plugin_savepoint($result, '2015060116.03', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060118) {
        $table = new xmldb_table('local_o365_appassign');
        $index = new xmldb_index('userobjectid', XMLDB_INDEX_UNIQUE, array('userobjectid'));
        if ($dbman->index_exists($table, $index)) {
            $dbman->drop_index($table, $index);
        }
        $field = new xmldb_field('userobjectid', XMLDB_TYPE_CHAR, '255', null, null, null, null, 'muserid');
        if ($dbman->field_exists($table, $field)) {
            $dbman->drop_field($table, $field);
        }
        $field = new xmldb_field('photoid', XMLDB_TYPE_CHAR, '255', null, null, null, null, 'assigned');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }
        $field = new xmldb_field('photoupdated', XMLDB_TYPE_INTEGER, '10', null, null, null, null, 'photoid');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }
        upgrade_plugin_savepoint($result, '2015060118', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060120) {
        // Delete custom profile fields for data type o365 and oidc which are no longer used.
        $fields = $DB->get_records_sql("SELECT * FROM {user_info_field} WHERE datatype IN ('o365', 'oidc')");
        foreach ($fields as $field) {
            $DB->delete_records('user_info_data', array('fieldid' => $field->id));
            $DB->delete_records('user_info_field', array('id' => $field->id));
        }
        upgrade_plugin_savepoint($result, '2015060120', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060126.01) {
        $table = new xmldb_table('local_o365_appassign');
        $field = new xmldb_field('photoupdated', XMLDB_TYPE_INTEGER, '10', null, null, null, null, 'photoid');
        if ($dbman->field_exists($table, $field)) {
            $field->setNotNull(FALSE);
            $dbman->change_field_notnull($table, $field);
        }
        upgrade_plugin_savepoint($result, '2015060126.01', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060128) {
        if (!$dbman->table_exists('local_o365_coursegroupdata')) {
            $dbman->install_one_table_from_xmldb_file(__DIR__.'/install.xml', 'local_o365_coursegroupdata');
        }
        upgrade_plugin_savepoint($result, '2015060128', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060128.02) {
        $config = get_config('local_o365');
        if (!empty($config->creategroups)) {
            set_config('creategroups', 'onall', 'local_o365');
        }
        upgrade_plugin_savepoint($result, '2015060128.02', 'local', 'o365');
    }

    if ($oldversion < 2015060129.02) {
        // Define field openidconnect to be added to local_o365_matchqueue.
        $table = new xmldb_table('local_o365_matchqueue');
        $field = new xmldb_field('openidconnect', XMLDB_TYPE_INTEGER, '1', null, XMLDB_NOTNULL, null, '0', 'o365username');

        // Conditionally launch add field openidconnect.
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }

        // O365 savepoint reached.
        upgrade_plugin_savepoint(true, 2015060129.02, 'local', 'o365');
    }

    if ($result && $oldversion < 2015060129.03) {
        // Drop index.
        $table = new xmldb_table('local_o365_token');
        $index = new xmldb_index('user', XMLDB_INDEX_NOTUNIQUE, ['user_id']);
        if ($dbman->index_exists($table, $index)) {
            $dbman->drop_index($table, $index);
        }

        $table = new xmldb_table('local_o365_calsub');
        $index = new xmldb_index('user', XMLDB_INDEX_NOTUNIQUE, ['user_id']);
        if ($dbman->index_exists($table, $index)) {
            $dbman->drop_index($table, $index);
        }

        upgrade_plugin_savepoint($result, '2015060129.03', 'local', 'o365');
    }

    if ($result && $oldversion < 2015060132) {
        $table = new xmldb_table('local_o365_calsub');
        $field = new xmldb_field('o365calemail', XMLDB_TYPE_CHAR, '255', null, null, null, null, 'o365calid');
        if (!$dbman->field_exists($table, $field)) {
            $dbman->add_field($table, $field);
        }
        upgrade_plugin_savepoint($result, '2015060132', 'local', 'o365');
    }

    return $result;
}
