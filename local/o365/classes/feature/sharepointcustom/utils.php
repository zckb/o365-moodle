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

namespace local_o365\feature\sharepointcustom;

class utils {
    /**
     * Determine whether course usergroups are enabled or not.
     *
     * @return bool True if group creation is enabled. False otherwise.
     */
    public static function is_enabled() {
        $creategroups = get_config('local_o365', 'creategroups');
        return ($creategroups === 'oncustom' || $creategroups === 'onall') ? true : false;
    }

    /**
     * Get an array of enabled courses.
     *
     * @return array Array of course IDs, or TRUE if all courses enabled.
     */
    public static function get_enabled_courses() {
        $creategroups = get_config('local_o365', 'creategroups');
        if ($creategroups === 'onall') {
            return true;
        } else if ($creategroups === 'oncustom') {
            $coursesenabled = get_config('local_o365', 'usergroupcustom');
            $coursesenabled = @json_decode($coursesenabled, true);
            if (!empty($coursesenabled) && is_array($coursesenabled)) {
                return array_keys($coursesenabled);
            }
        }
        return [];
    }

    /**
     * Enable or disable course subsite.
     *
     * @param int $courseid The ID of the course.
     * @param bool $enabled Whether to enable or disable.
     */
    public static function set_course_subsite_enabled($courseid, $enabled = true) {
        // Check that custom subsites is enabled.
        $customsubsitesenabled = get_config('local_o365', 'sharepointcourseselect');
        if ($customsubsitesenabled === 'oncustom') {
            $customsubsitesconfig = get_config('local_o365', 'sharepointsubsitescustom');
            $customsubsitesconfig = @json_decode($customsubsitesconfig, true);
            if (empty($customsubsitesconfig) || !is_array($customsubsitesconfig)) {
                $customsubsitesconfig = [];
            }
            if ($enabled === true) {
                $customsubsitesconfig[$courseid] = $enabled;
            } else {
                if (isset($customsubsitesconfig[$courseid])) {
                    unset($customsubsitesconfig[$courseid]);
                }
            }
            set_config('sharepointsubsitescustom', json_encode($customsubsitesconfig), 'local_o365');
            // Create the course subsite.
            // $createsubsite = {};
            // $createsubsite->create_course_site($courseid);
        }
    }

   /**
     * Determine whether a course subsite is enabled or disabled.
     *
     * @param int $courseid The ID of the course.
     * @param string $feature The feature to check.
     * @return bool Whether the feature is enabled or not.
     */
    public static function course_is_sharepoint_enabled($courseid) {
        $customsubsitesenabled = get_config('local_o365', 'sharepointcourseselect');
        if ($customsubsitesenabled === 'off') {
            return true;
        } else if ($customsubsitesenabled === 'oncustom') {
            $config = get_config('local_o365', 'sharepointsubsitescustom');
            $config = @json_decode($config, true);
            return (!empty($config) && is_array($config) && isset($config[$courseid]))
                ? true : false;
        }
        return false;
    }

    /**
     * Determine whether or not a subsite can be created for a course.
     *
     * @param string $course A course record to create the subsite from.
     * @return bool True if course subsite can be created, false otherwise.
     */
    public static function course_subsite_enabled($course) {
        $customsubsitesenabled = get_config('local_o365', 'sharepointcourseselect');
        $subsitesconfig = get_config('local_o365', 'sharepointsubsitescustom');
        $subsitesconfig = json_decode($subsitesconfig, true);
        $courseid = $course->id;
        $courseinconfig = array_key_exists($courseid, $subsitesconfig);// empty($subsitesconfig->$courseid);
        $courseconfigval = false;
        if ($courseinconfig) {
            $courseconfigval = $subsitesconfig[$courseid];
        }
        if ($customsubsitesenabled === 'off' || ($customsubsitesenabled === 'oncustom' && $courseinconfig == true && $courseconfigval == true)) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * Change whether group features are enabled for a course.
     *
     * @param int $courseid The ID of the course.
     * @param array $features Array of features to enable or disable.
     * @param bool $enabled Whether to enable or disable.
     */
    // public static function set_course_group_feature_enabled($courseid, array $features, $enabled = true) {
    //     $usergroupconfig = get_config('local_o365', 'usergroupcustomfeatures');
    //     $usergroupconfig = @json_decode($usergroupconfig, true);
    //     if (empty($usergroupconfig) || !is_array($usergroupconfig)) {
    //         $usergroupconfig = [];
    //     }
    //     if (!isset($usergroupconfig[$courseid])) {
    //         $usergroupconfig[$courseid] = [];
    //     }
    //     if ($enabled === true) {
    //         foreach ($features as $feature) {
    //             $usergroupconfig[$courseid][$feature] = $enabled;
    //         }
    //     } else {
    //         foreach ($features as $feature) {
    //             if (isset($usergroupconfig[$courseid][$feature])) {
    //                 unset($usergroupconfig[$courseid][$feature]);
    //             }
    //         }
    //     }
    //     set_config('usergroupcustomfeatures', json_encode($usergroupconfig), 'local_o365');
    // }

    /**
     * Enable or disable a feature for all group courses.
     *
     * @param string $feature The feature to enable or disable.
     * @param bool $enabled Whether to enable or disable.
     */
    // public static function bulk_set_group_feature_enabled($feature, $enabled) {
    //     $usergroupconfig = get_config('local_o365', 'usergroupcustomfeatures');
    //     $usergroupconfig = @json_decode($usergroupconfig, true);
    //     if ($enabled === true) {
    //         if (empty($usergroupconfig) || !is_array($usergroupconfig)) {
    //             $usergroupconfig = [];
    //         }
    //         $enabledcourses = static::get_enabled_courses();
    //         foreach ($enabledcourses as $courseid) {
    //             if (!isset($usergroupconfig[$courseid])) {
    //                 $usergroupconfig[$courseid] = [];
    //             }
    //             $usergroupconfig[$courseid][$feature] = true;
    //         }
    //     } else {
    //         if (empty($usergroupconfig) || !is_array($usergroupconfig)) {
    //             return true;
    //         } else {
    //             foreach ($usergroupconfig as $courseid => $features) {
    //                 if (isset($features[$feature])) {
    //                     unset($usergroupconfig[$courseid][$feature]);
    //                 }
    //             }
    //         }
    //     }
    //     set_config('usergroupcustomfeatures', json_encode($usergroupconfig), 'local_o365');
    // }
}
