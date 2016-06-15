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
 * @package repository_office365
 * @author James McQuillan <james.mcquillan@remote-learner.net>
 * @license http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 * @copyright (C) 2014 onwards Microsoft, Inc. (http://microsoft.com/)
 */

/**
 * Office 365 repository.
 */
class repository_office365 extends \repository {

    /** @var \local_o365\httpclient An HTTP client to use. */
    protected $httpclient;

    /** @var bool Whether onedrive is configured. */
    protected $onedriveconfigured = false;

    /** @var bool Whether sharepoint is configured. */
    protected $sharepointconfigured = false;

    /** @var bool Whether the Microsoft Graph API is configured. */
    protected $unifiedconfigured = false;

    /** @var \local_o365\oauth2\clientdata A clientdata object to use with an o365 api class. */
    protected $clientdata = null;

    /**
     * Constructor
     *
     * @param int $repositoryid repository instance id
     * @param int|stdClass $context a context id or context object
     * @param array $options repository options
     * @param int $readonly indicate this repo is readonly or not
     */
    public function __construct($repositoryid, $context = SYSCONTEXTID, $options = array(), $readonly = 0) {
        parent::__construct($repositoryid, $context, $options, $readonly);
        $this->httpclient = new \local_o365\httpclient();
        if (\local_o365\utils::is_configured()) {
            $this->clientdata = \local_o365\oauth2\clientdata::instance_from_oidc();
        }
        $this->onedriveconfigured = \local_o365\rest\onedrive::is_configured();
        $this->unifiedconfigured = \local_o365\rest\unified::is_configured();
        $this->sharepointconfigured = \local_o365\rest\sharepoint::is_configured();
    }

    /**
     * Get a Microsoft Graph API token.
     *
     * @param bool $system If true, get a system API ser token instead of the user's token.
     * @param int|null $userid The userid to get a token for. If null, the current user will be used.
     * @return \local_o365\oauth2\token A Microsoft Graph API token object.
     */
    protected function get_unified_token($system = false, $userid = null) {
        global $USER;
        $resource = \local_o365\rest\unified::get_resource();
        if ($system === true) {
            return \local_o365\oauth2\systemtoken::instance(null, $resource, $this->clientdata, $this->httpclient);
        } else {
            $userid = (!empty($userid)) ? $userid : $USER->id;
            return \local_o365\oauth2\token::instance($userid, $resource, $this->clientdata, $this->httpclient);
        }
    }

    /**
     * Get a OneDrive token.
     *
     * @param bool $system If true, get a system API ser token instead of the user's token.
     * @param int|null $userid The userid to get a token for. If null, the current user will be used.
     * @return \local_o365\oauth2\token A OneDrive token object.
     */
    protected function get_onedrive_token($system = false, $userid = null) {
        global $USER;
        $resource = \local_o365\rest\onedrive::get_resource();
        if ($system === true) {
            return \local_o365\oauth2\systemtoken::instance(null, $resource, $this->clientdata, $this->httpclient);
        } else {
            $userid = (!empty($userid)) ? $userid : $USER->id;
            return \local_o365\oauth2\token::instance($userid, $resource, $this->clientdata, $this->httpclient);
        }
    }

    /**
     * Get a SharePoint token.
     *
     * @param bool $system If true, get a system API ser token instead of the user's token.
     * @param int|null $userid The userid to get a token for. If null, the current user will be used.
     * @return \local_o365\oauth2\token A SharePoint token object.
     */
    protected function get_sharepoint_token($system = false, $userid = null) {
        global $USER;
        $resource = \local_o365\rest\sharepoint::get_resource();
        if ($system === true) {
            return \local_o365\oauth2\systemtoken::instance(null, $resource, $this->clientdata, $this->httpclient);
        } else {
            $userid = (!empty($userid)) ? $userid : $USER->id;
            return \local_o365\oauth2\token::instance($userid, $resource, $this->clientdata, $this->httpclient);
        }
    }

    /**
     * Get a Microsoft Graph API client.
     *
     * @param bool $system If true, get a system API ser token instead of the user's token.
     * @param int|null $userid The userid to get an API client for. If null, the current user will be used.
     * @return \local_o365\rest\unified A Microsoft Graph API client object.
     */
    protected function get_unified_apiclient($system = false, $userid = null) {
        if ($this->unifiedconfigured === true) {
            $token = $this->get_unified_token($system, $userid);
            if (!empty($token)) {
                return new \local_o365\rest\unified($token, $this->httpclient);
            }
        }
        return false;
    }

    /**
     * Get a onedrive API client.
     *
     * @param bool $system If true, get a system API ser token instead of the user's token.
     * @param int|null $userid The userid to get an API client for. If null, the current user will be used.
     * @return \local_o365\rest\onedrive A onedrive API client object.
     */
    protected function get_onedrive_apiclient($system = false, $userid = null) {
        if ($this->onedriveconfigured === true) {
            $token = $this->get_onedrive_token($system, $userid);
            if (!empty($token)) {
                return new \local_o365\rest\onedrive($token, $this->httpclient);
            }
        }
        return false;
    }

    /**
     * Get a sharepoint API client.
     *
     * @param bool $system If true, get a system API ser token instead of the user's token.
     * @param int|null $userid The userid to get an API client for. If null, the current user will be used.
     * @return \local_o365\rest\sharepoint A sharepoint API client object.
     */
    protected function get_sharepoint_apiclient($system = false, $userid = null) {
        if ($this->sharepointconfigured === true) {
            $token = $this->get_sharepoint_token($system, $userid);
            if (!empty($token)) {
                return new \local_o365\rest\sharepoint($token, $this->httpclient);
            }
        }
        return false;
    }

    /**
     * Given a path, and perhaps a search, get a list of files.
     *
     * See details on {@link http://docs.moodle.org/dev/Repository_plugins}
     *
     * @param string $path this parameter can a folder name, or a identification of folder
     * @param string $page the page number of file list
     * @return array the list of files, including meta infomation, containing the following keys
     *           manage, url to manage url
     *           client_id
     *           login, login form
     *           repo_id, active repository id
     *           login_btn_action, the login button action
     *           login_btn_label, the login button label
     *           total, number of results
     *           perpage, items per page
     *           page
     *           pages, total pages
     *           issearchresult, is it a search result?
     *           list, file list
     *           path, current path and parent path
     */
    public function get_listing($path = '', $page = '') {
        global $OUTPUT, $SESSION, $USER;

        if (\local_o365\utils::is_configured() !== true) {
            throw new \moodle_exception('errorauthoidcnotconfig', 'repository_office365');
        }

        $clientid = optional_param('client_id', '', PARAM_TEXT);
        if (!empty($clientid)) {
            $SESSION->repository_office365['curpath'][$clientid] = $path;
        }

        // If we were launched from a course context (or child of course context), initialize the file picker in the correct course.
        if (!empty($this->context)) {
            $context = $this->context->get_course_context(false);
        }
        if (empty($context)) {
            $context = \context_system::instance();
        }
        if ($this->sharepointconfigured === true && $context instanceof \context_course) {
            if (empty($path)) {
                $path = '/courses/'.$context->instanceid;
            }
        }

        $list = [];
        $breadcrumb = [['name' => $this->name, 'path' => '/']];

        $unifiedactive = false;
        if ($this->unifiedconfigured === true) {
            $unifiedtoken = $this->get_unified_token();
            if (!empty($unifiedtoken)) {
                $unifiedactive = true;
            }
        }

        $onedriveactive = false;
        if ($this->onedriveconfigured === true) {
            $onedrivetoken = $this->get_onedrive_token();
            if (!empty($onedrivetoken)) {
                $onedriveactive = true;
            }
        }

        $sharepointactive = false;
        if ($this->sharepointconfigured === true) {
            $sharepointtoken = $this->get_sharepoint_token();
            if (!empty($sharepointtoken)) {
                $sharepointactive = true;
            }
        }

        $courses = enrol_get_users_courses($USER->id, true);
        $showgroups = false;
        if (\local_o365\rest\unified::is_configured() === true) {
            foreach ($courses as $course) {
                if (\local_o365\feature\usergroups\utils::course_is_group_enabled($course->id)) {
                    if (\local_o365\feature\usergroups\utils::course_is_group_feature_enabled($course->id, 'onedrive')) {
                        $showgroups = true;
                        break;
                    }
                }
            }
        }

        if (strpos($path, '/my/') === 0) {
            if ($unifiedactive === true) {
                // Path is in my files.
                list($list, $breadcrumb) = $this->get_listing_my_unified(substr($path, 3));
            } else if ($onedriveactive === true) {
                // Path is in my files.
                list($list, $breadcrumb) = $this->get_listing_my(substr($path, 3));
            }
        } else if (strpos($path, '/courses/') === 0) {
            if ($sharepointactive === true) {
                // Path is in course files.
                list($list, $breadcrumb) = $this->get_listing_course(substr($path, 8));
            }
        } else if (strpos($path, '/groups/') === 0) {
            if ($showgroups === true) {
                // Path is in group files.
                list($list, $breadcrumb) = $this->get_listing_groups(substr($path, 7));
            }
        } else if (strpos($path, '/trending/') === 0) {
            if ($unifiedactive === true) {
                // Path is in trending files.
                list($list, $breadcrumb) = $this->get_listing_trending_unified(substr($path, 9));
            }
        } else if (strpos($path, '/office365video/') === 0) {
            if ($sharepointactive === true) {
                // Path is in office365 channels/videos.
                list($list, $breadcrumb) = $this->get_listing_videos(substr($path, 15));
            }
        } else {
            if ($unifiedactive === true || $onedriveactive === true) {
                $list[] = [
                    'title' => get_string('myfiles', 'repository_office365'),
                    'path' => '/my/',
                    'thumbnail' => $OUTPUT->pix_url('onedrive', 'repository_office365')->out(false),
                    'children' => [],
                ];
            }
            if ($sharepointactive === true) {
                $list[] = [
                    'title' => get_string('courses', 'repository_office365'),
                    'path' => '/courses/',
                    'thumbnail' => $OUTPUT->pix_url('sharepoint', 'repository_office365')->out(false),
                    'children' => [],
                ];
                $sharepoint = $this->get_sharepoint_apiclient();
                // Retrieve api url for video service.
                $url = $sharepoint->videoservice_discover();
                if (!empty($url)) {
                    $list[] = [
                        'title' => get_string('office365video', 'repository_office365'),
                        'path' => '/office365video/',
                        'thumbnail' => $OUTPUT->pix_url('office365video', 'repository_office365')->out(false),
                        'children' => [],
                    ];
                }
            }
            if ($showgroups === true) {
                $list[] = [
                    'title' => get_string('groups', 'repository_office365'),
                    'path' => '/groups/',
                    'thumbnail' => $OUTPUT->pix_url('coursegroups', 'repository_office365')->out(false),
                    'children' => [],
                ];
            }
            if ($unifiedactive === true) {
                $list[] = [
                    'title' => get_string('trendingaround', 'repository_office365'),
                    'path' => '/trending/',
                    'thumbnail' => $OUTPUT->pix_url('delve', 'repository_office365')->out(false),
                    'children' => [],
                ];
            }
        }
        if ($this->path_is_upload($path) === true) {
            return [
                'dynload' => true,
                'nologin' => true,
                'nosearch' => true,
                'path' => $breadcrumb,
                'upload' => [
                    'label' => get_string('file', 'repository_office365'),
                ],
            ];
        }

        return [
            'dynload' => true,
            'nologin' => true,
            'nosearch' => true,
            'list' => $list,
            'path' => $breadcrumb,
        ];
    }

    /**
     * Determine whether a given path is an upload path.
     *
     * @param string $path A path to check.
     * @return bool Whether the path is an upload path.
     */
    protected function path_is_upload($path) {
        return (substr($path, -strlen('/upload/')) === '/upload/') ? true : false;
    }

    /**
     * Process uploaded file.
     *
     * @return array Array of uploaded file information.
     */
    public function upload($saveasfilename, $maxbytes) {
        global $CFG, $USER, $SESSION;
        $caller = '\repository_office365::upload';

        $types = optional_param_array('accepted_types', '*', PARAM_RAW);
        $savepath = optional_param('savepath', '/', PARAM_PATH);
        $itemid = optional_param('itemid', 0, PARAM_INT);
        $license = optional_param('license', $CFG->sitedefaultlicense, PARAM_TEXT);
        $author = optional_param('author', '', PARAM_TEXT);
        $areamaxbytes = optional_param('areamaxbytes', FILE_AREA_MAX_BYTES_UNLIMITED, PARAM_INT);
        $overwriteexisting = optional_param('overwrite', false, PARAM_BOOL);
        $clientid = optional_param('client_id', '', PARAM_TEXT);

        $filepath = '/';
        if (!empty($SESSION->repository_office365)) {
            if (isset($SESSION->repository_office365['curpath']) && isset($SESSION->repository_office365['curpath'][$clientid])) {
                $filepath = $SESSION->repository_office365['curpath'][$clientid];
                if (strpos($filepath, '/my/') === 0) {
                    $clienttype = 'onedrive';
                    $filepath = substr($filepath, 3);
                } else if (strpos($filepath, '/courses/') === 0) {
                    $clienttype = 'sharepoint';
                    $filepath = substr($filepath, 8);
                } else if (strpos($filepath, '/office365video/') === 0) {
                    $clienttype = 'office365video';
                    $filepath = substr($filepath, 15);
                } else {
                    $errmsg = get_string('errorbadclienttype', 'repository_office365');
                    $debugdata = [
                        'filepath' => $filepath,
                    ];
                    \local_o365\utils::debug($errmsg, $caller, $debugdata);
                    throw new \moodle_exception('errorbadclienttype', 'repository_office365');
                }
            }
        }
        if ($this->path_is_upload($filepath) === true) {
            $filepath = substr($filepath, 0, -strlen('/upload/'));
        }
        $filename = (!empty($saveasfilename)) ? $saveasfilename : $_FILES['repo_upload_file']['name'];
        $filename = clean_param($filename, PARAM_FILE);
        $content = file_get_contents($_FILES['repo_upload_file']['tmp_name']);

        if ($clienttype === 'onedrive') {
            if ($this->unifiedconfigured === true) {
                $apiclient = $this->get_unified_apiclient();
                $parentid = (!empty($filepath)) ? substr($filepath, 1) : '';
                $result = $apiclient->create_file($parentid, $filename, $content, 'application/octet-stream');
            } else {
                $apiclient = $this->get_onedrive_apiclient();
                $result = $apiclient->create_file($filepath, $filename, $content);
            }
            $source = $this->pack_reference(['id' => $result['id'], 'source' => 'onedrive']);
        } else if ($clienttype === 'sharepoint') {
            $pathtrimmed = trim($filepath, '/');
            $pathparts = explode('/', $pathtrimmed);
            if (!is_numeric($pathparts[0])) {
                $errmsg = get_string('errorbadpath', 'repository_office365');
                $debugdata = [
                    'filepath' => $filepath,
                ];
                \local_o365\utils::debug($errmsg, $caller, $debugdata);
                throw new \moodle_exception('errorbadpath', 'repository_office365');
            }
            $courseid = (int)$pathparts[0];
            unset($pathparts[0]);
            $relpath = (!empty($pathparts)) ? implode('/', $pathparts) : '';
            $fullpath = (!empty($relpath)) ? '/'.$relpath : '/';
            $courses = enrol_get_users_courses($USER->id);
            if (!isset($courses[$courseid])) {
                $errmsg = get_string('erroraccessdenied', 'repository_office365');
                $debugdata = [
                    'courseid' => $courseid,
                    'courses' => $courses,
                    'filepath' => $filepath,
                ];
                \local_o365\utils::debug($errmsg, $caller, $debugdata);
                throw new \moodle_exception('erroraccessdenied', 'repository_office365');
            }
            $curcourse = $courses[$courseid];
            unset($courses);
            $sharepoint = $this->get_sharepoint_apiclient();
            $parentsiteuri = $sharepoint->get_course_subsite_uri($curcourse->id);
            $sharepoint->set_site($parentsiteuri);
            $result = $sharepoint->create_file($fullpath, $filename, $content);
            $source = $this->pack_reference(['id' => $result['id'], 'source' => $clienttype, 'parentsiteuri' => $parentsiteuri]);
        } else if ($clienttype === 'office365video') {
            if ($this->sharepointconfigured == true) {
                $sharepoint = $this->get_sharepoint_apiclient();
                // Retrieve api url for video service.
                $url = $sharepoint->videoservice_discover();
                if (!empty($url)) {
                    $sharepoint->override_resource($url);
                    $parentid = (!empty($filepath)) ? substr($filepath, 1) : '';
                    $videoobject = $sharepoint->create_video_placeholder($parentid, '', $filename, $filename);
                    if (!empty($videoobject)) {
                        $result = $sharepoint->upload_video($videoobject['ChannelID'], $videoobject['ID'], $_FILES['repo_upload_file']['tmp_name']);
                        $parseurl = explode('/', $videoobject['Url']);
                        $downloadurl = "https://".$parseurl[2]."/_api/SP.AppContextSite(@target)/Web/"."GetFileByServerRelativeUrl('".$videoobject['ServerRelativeUrl'].
                                "')/"."$"."value?@target='https://".$parseurl[2]."/portals/".$parseurl[4]."'";
                        $url = "https://".$parseurl[2]."/portals/hub/_layouts/15/PointPublishing.aspx?app=video&"."p=p&chid=".$videoobject['ChannelID']."&vid=".$videoobject['ID'];
                        $source = $this->pack_reference(['id' => $videoobject['odata.id'],
                            'source' => 'office365video',
                            'url' => $url,
                            'downloadurl' => $downloadurl]);
                    }
                }
            }
        } else {
            $errmsg = get_string('errorbadclienttype', 'repository_office365');
            $debugdata = [
                'clienttype' => $clienttype,
            ];
            \local_o365\utils::debug($errmsg, $caller, $debugdata);
            throw new \moodle_exception('errorbadclienttype', 'repository_office365');
        }

        $downloadedfile = $this->get_file($source, $filename);
        $record = new \stdClass;
        $record->filename = $filename;
        $record->filepath = $savepath;
        $record->component = 'user';
        $record->filearea = 'draft';
        $record->itemid = $itemid;
        $record->license = $license;
        $record->author = $author;
        $usercontext = \context_user::instance($USER->id);
        $now = time();
        $record->contextid = $usercontext->id;
        $record->timecreated = $now;
        $record->timemodified = $now;
        $record->userid = $USER->id;
        $record->sortorder = 0;
        $record->source = $this->build_source_field($source);
        $info = \repository::move_to_filepool($downloadedfile['path'], $record);
        if ($clienttype === 'office365video') {
            $info['url'] = $url;
        }
        return $info;
    }

    /**
     * Get listing for a group folder.
     *
     * @param string $path Folder path.
     * @return array List of $list array and $path array.
     */
    protected function get_listing_groups($path = '') {
        global $OUTPUT, $USER, $DB;
        $caller = '\repository_office365::get_listing_groups';
        $list = [];
        $breadcrumb = [
            ['name' => $this->name, 'path' => '/'],
            ['name' => get_string('groups', 'repository_office365'), 'path' => '/groups/'],
        ];

        $coursesbyid = enrol_get_users_courses($USER->id, true);

        if ($path === '/') {
            // Show available courses.
            $showgroups = false;
            $enabledcourses = \local_o365\feature\usergroups\utils::get_enabled_courses_with_feature('onedrive');
            foreach ($coursesbyid as $course) {
                if ($enabledcourses === true || in_array($course->id, $enabledcourses)) {
                    $list[] = [
                        'title' => $course->shortname,
                        'path' => '/groups/'.$course->id,
                        'thumbnail' => $OUTPUT->pix_url(file_folder_icon(90))->out(false),
                        'children' => [],
                    ];
                }
            }
        } else {
            $pathtrimmed = trim($path, '/');
            $pathparts = explode('/', $pathtrimmed);
            if (!is_numeric($pathparts[0]) || !isset($coursesbyid[$pathparts[0]])
                    || \local_o365\feature\usergroups\utils::course_is_group_enabled($pathparts[0]) !== true
                    || \local_o365\feature\usergroups\utils::course_is_group_feature_enabled($pathparts[0], 'onedrive') !== true) {
                \local_o365\utils::debug(get_string('errorbadpath', 'repository_office365'), $caller, ['path' => $path]);
                throw new \moodle_exception('errorbadpath', 'repository_office365');
            }
            $courseid = (int)$pathparts[0];
            $curpath = '/groups/'.$courseid;
            $breadcrumb[] = ['name' => $coursesbyid[$courseid]->shortname, 'path' => $curpath];

            $sql = 'SELECT g.*
                      FROM {groups} g
                      JOIN {groups_members} m ON m.groupid = g.id
                     WHERE m.userid = ? AND g.courseid = ?';
            $coursegroups = $DB->get_records_sql($sql, [$USER->id, $courseid]);

            if (count($pathparts) === 1) {
                $list[] = [
                    'title' => get_string('defaultgroupsfolder', 'repository_office365'),
                    'path' => $curpath.'/coursegroup/',
                    'thumbnail' => $OUTPUT->pix_url(file_folder_icon(90))->out(false),
                    'children' => [],
                ];

                foreach ($coursegroups as $group) {
                    $list[] = [
                        'title' => $group->name,
                        'path' => $curpath.'/'.$group->id.'/',
                        'thumbnail' => $OUTPUT->pix_url(file_folder_icon(90))->out(false),
                        'children' => [],
                    ];
                }
            } else {
                // Validate the received group identifier.
                if (!is_numeric($pathparts[1]) && $pathparts[1] !== 'coursegroup') {
                    \local_o365\utils::debug(get_string('errorbadpath', 'repository_office365'), $caller, ['path' => $path]);
                    throw new \moodle_exception('errorbadpath', 'repository_office365');
                }
                $curpath .= '/'.$pathparts[1].'/';
                if ($pathparts[1] === 'coursegroup') {
                    $breadcrumb[] = ['name' => get_string('defaultgroupsfolder', 'repository_office365'), 'path' => $curpath];
                    $filters = ['type' => 'group', 'subtype' => 'course', 'moodleid' => $courseid];
                    $group = $DB->get_record('local_o365_objects', $filters);
                } else {
                    // Validate the user is a member of the group.
                    if (!isset($coursegroups[$pathparts[1]])) {
                        \local_o365\utils::debug(get_string('errorbadpath', 'repository_office365'), $caller, ['path' => $path]);
                        throw new \moodle_exception('errorbadpath', 'repository_office365');
                    }
                    $groupid = (int)$pathparts[1];
                    $group = $DB->get_record('groups', ['id' => $groupid]);
                    $breadcrumb[] = ['name' => $group->name, 'path' => $curpath];
                    $filters = ['type' => 'group', 'subtype' => 'usergroup', 'moodleid' => $groupid];
                    $group = $DB->get_record('local_o365_objects', $filters);
                }

                $intragrouppath = $pathparts;
                unset($intragrouppath[0], $intragrouppath[1]);
                $curparent = trim(end($intragrouppath));

                if (!empty($group)) {
                    $unified = $this->get_unified_apiclient();

                    if (!empty($curparent)) {
                        $metadata = $unified->get_group_file_metadata($group->objectid, $curparent);
                        if (!empty($metadata['parentReference']) && !empty($metadata['parentReference']['path'])) {
                            $parentrefpath = substr($metadata['parentReference']['path'], (strpos($metadata['parentReference']['path'], ':') + 1));
                            $cache = \cache::make('repository_office365', 'unifiedgroupfolderids');
                            $result = $cache->set($parentrefpath.'/'.$metadata['name'], $metadata['id']);
                            if (!empty($parentrefpath)) {
                                $parentrefpath = explode('/', trim($parentrefpath, '/'));
                                $currentfullpath = '';
                                foreach ($parentrefpath as $folder) {
                                    $currentfullpath .= '/'.$folder;
                                    $folderid = $cache->get($currentfullpath);
                                    $breadcrumb[] = ['name' => $folder, 'path' => $curpath.$folderid];
                                }
                            }
                        }
                        $breadcrumb[] = ['name' => $metadata['name'], 'path' => $curpath.$metadata['id']];
                    }
                    $contents = $unified->get_group_files($group->objectid, $curparent);
                    $list = $this->contents_api_response_to_list($contents, $path, 'unifiedgroup', $group->objectid, false);
                } else {
                    \local_o365\utils::debug('Could not file group object record', $caller, ['path' => $path]);
                    $list = [];
                }
            }
        }

        return [$list, $breadcrumb];
    }

    /**
     * Get listing for a course folder.
     *
     * @param string $path Folder path.
     * @return array List of $list array and $path array.
     */
    protected function get_listing_course($path = '') {
        global $USER, $OUTPUT;

        $caller = '\repository_office365::get_listing_course';
        $list = [];
        $breadcrumb = [
            ['name' => $this->name, 'path' => '/'],
            ['name' => get_string('courses', 'repository_office365'), 'path' => '/courses/'],
        ];

        $reqcap = \local_o365\rest\sharepoint::get_course_site_required_capability();
        $courses = get_user_capability_course($reqcap, $USER->id, true, 'shortname');
        // Reindex courses array using course id.
        $coursesbyid = [];
        foreach ($courses as $i => $course) {
            $coursesbyid[$course->id] = $course;
            unset($courses[$i]);
        }
        unset($courses);

        if ($path === '/') {
            // Show available courses.
            foreach ($coursesbyid as $course) {
                $list[] = [
                    'title' => $course->shortname,
                    'path' => '/courses/'.$course->id,
                    'thumbnail' => $OUTPUT->pix_url(file_folder_icon(90))->out(false),
                    'children' => [],
                ];
            }
        } else {
            $pathtrimmed = trim($path, '/');
            $pathparts = explode('/', $pathtrimmed);
            if (!is_numeric($pathparts[0])) {
                $errmsg = get_string('errorbadpath', 'repository_office365');
                $debugdata = [
                    'path' => $path,
                ];
                \local_o365\utils::debug($errmsg, $caller, $debugdata);
                throw new \moodle_exception('errorbadpath', 'repository_office365');
            }
            $courseid = (int)$pathparts[0];
            unset($pathparts[0]);
            $relpath = (!empty($pathparts)) ? implode('/', $pathparts) : '';
            if (isset($coursesbyid[$courseid])) {
                if ($this->path_is_upload($path) === false) {
                    $sharepointclient = $this->get_sharepoint_apiclient();
                    if (!empty($sharepointclient)) {
                        $parentsiteuri = $sharepointclient->get_course_subsite_uri($coursesbyid[$courseid]->id);
                        $sharepointclient->set_site($parentsiteuri);
                        try {
                            $fullpath = (!empty($relpath)) ? '/'.$relpath : '/';
                            $contents = $sharepointclient->get_files($fullpath);
                            $list = $this->contents_api_response_to_list($contents, $path, 'sharepoint', $parentsiteuri);
                        } catch (\Exception $e) {
                            $errmsg = 'Exception when retrieving share point files';
                            $debugdata = [
                                'fullpath' => (!empty($relpath)) ? '/'.$relpath : '/',
                                'message' => $e->getMessage(),
                            ];
                            \local_o365\utils::debug($errmsg, $caller, $debugdata);
                            $list = [];
                        }
                    }
                }

                $curpath = '/courses/'.$courseid;
                $breadcrumb[] = ['name' => $coursesbyid[$courseid]->shortname, 'path' => $curpath];
                foreach ($pathparts as $i => $pathpart) {
                    if (!empty($pathpart)) {
                        $curpath .= '/'.$pathpart;
                        if ($i === (count($pathparts)) && $pathpart === 'upload') {
                            $pathpart = get_string('upload', 'repository_office365');
                        }
                        $breadcrumb[] = ['name' => $pathpart, 'path' => $curpath];
                    }
                }
            }
        }

        return [$list, $breadcrumb];
    }

    /**
     * Get listing for a o365 video folder.
     *
     * @param string $path Folder path.
     * @return array List of $list array and $path array.
     */
    protected function get_listing_videos($path = '') {
        $path = (empty($path)) ? '/' : $path;
        $list = [];
        $sharepoint = $this->get_sharepoint_apiclient();
        // Retrieve api url for video service.
        $url = $sharepoint->videoservice_discover();
        if (!empty($url)) {
            $sharepoint->override_resource($url);
            if ($this->path_is_upload($path) === true) {
                $path = substr($path, 0, -strlen('/upload/'));
            } else if ($path === '/') {
                // Get list of Channels.
                $contents = $sharepoint->get_video_channels();
                $list = $this->contents_api_response_to_list($contents, $path, 'office365video', null, false);
            } else {
                // Get videos of a Channel.
                $contents = $sharepoint->get_all_channel_videos(substr($path, 1));
                $list = $this->contents_api_response_to_list($contents, $path, 'office365video', null, true);
            }
            if ($path !== '/') {
                $channel = $sharepoint->get_video_channel(substr($path, 1));
            }
        }
        // Generate path.
        $breadcrumb = [
            ['name' => $this->name, 'path' => '/'],
            ['name' => get_string('office365video', 'repository_office365'), 'path' => '/office365video/'],
        ];
        $pathparts = explode('/', $path);
        $curpath = '/office365video';
        // Remove empty paths (we do this in a separate loop for proper upload detection in the next loop.
        foreach ($pathparts as $i => $pathpart) {
            if (empty($pathpart)) {
                unset($pathparts[$i]);
            }
        }
        $pathparts = array_values($pathparts);
        if (!empty($channel)) {
            array_push($breadcrumb, ['name' => $channel['Title'], 'path' => $curpath.'/'.$pathparts[0]]);
            array_splice($pathparts, 0, 1);
        }
        foreach ($pathparts as $i => $pathpart) {
            $curpath .= '/'.$pathpart;
            $pathname = $pathpart;
            if ($i === (count($pathparts) - 1) && $pathpart === 'upload') {
                $pathname = get_string('upload', 'repository_office365');
            }
            $breadcrumb[] = ['name' => $pathname, 'path' => $curpath];
        }
        return [$list, $breadcrumb];
    }

    /**
     * Get listing for a personal onedrive folder using the Microsoft Graph API.
     *
     * @param string $path Folder path.
     * @return array List of $list array and $path array.
     */
    protected function get_listing_my_unified($path = '') {
        $path = (empty($path)) ? '/' : $path;

        $list = [];

        $unified = $this->get_unified_apiclient();
        $realpath = $path;
        if ($this->path_is_upload($path) === true) {
            $realpath = substr($path, 0, -strlen('/upload/'));
        } else {
            $contents = $unified->get_files($realpath);
            $list = $this->contents_api_response_to_list($contents, $realpath, 'unified');
        }

        // Generate path.
        $strmyfiles = get_string('myfiles', 'repository_office365');
        $breadcrumb = [['name' => $this->name, 'path' => '/'], ['name' => $strmyfiles, 'path' => '/my/']];

        if ($realpath !== '/') {
            $metadata = $unified->get_file_metadata($realpath);
            if (!empty($metadata['parentReference']) && !empty($metadata['parentReference']['path'])) {
                $parentrefpath = substr($metadata['parentReference']['path'], (strpos($metadata['parentReference']['path'], ':') + 1));
                $cache = \cache::make('repository_office365', 'unifiedfolderids');
                $result = $cache->set($parentrefpath.'/'.$metadata['name'], $metadata['id']);
                if (!empty($parentrefpath)) {
                    $parentrefpath = explode('/', trim($parentrefpath, '/'));
                    $currentfullpath = '';
                    foreach ($parentrefpath as $folder) {
                        $currentfullpath .= '/'.$folder;
                        $folderid = $cache->get($currentfullpath);
                        $breadcrumb[] = ['name' => $folder, 'path' => '/my/'.$folderid];
                    }
                }
            }
            $breadcrumb[] = ['name' => $metadata['name'], 'path' => '/my/'.$metadata['id']];
        }

        if ($this->path_is_upload($path) === true) {
            $breadcrumb[] = ['name' => get_string('upload', 'repository_office365'), 'path' => '/my/'.$metadata['id'].'/upload/'];
        }

        return [$list, $breadcrumb];
    }

    /**
     * Get listing for a personal onedrive folder.
     *
     * @param string $path Folder path.
     * @return array List of $list array and $path array.
     */
    protected function get_listing_my($path = '') {
        $path = (empty($path)) ? '/' : $path;

        $list = [];
        if ($this->path_is_upload($path) !== true) {
            $onedrive = $this->get_onedrive_apiclient();
            $contents = $onedrive->get_contents($path);
            $list = $this->contents_api_response_to_list($contents, $path, 'onedrive');
        } else {
            $list = [];
        }

        // Generate path.
        $strmyfiles = get_string('myfiles', 'repository_office365');
        $breadcrumb = [['name' => $this->name, 'path' => '/'], ['name' => $strmyfiles, 'path' => '/my/']];
        $pathparts = explode('/', $path);
        $curpath = '/my';

        // Remove empty paths (we do this in a separate loop for proper upload detection in the next loop.
        foreach ($pathparts as $i => $pathpart) {
            if (empty($pathpart)) {
                unset($pathparts[$i]);
            }
        }
        $pathparts = array_values($pathparts);

        foreach ($pathparts as $i => $pathpart) {
            $curpath .= '/'.$pathpart;
            $pathname = $pathpart;
            if ($i === (count($pathparts) - 1) && $pathpart === 'upload') {
                $pathname = get_string('upload', 'repository_office365');
            }
            $breadcrumb[] = ['name' => $pathname, 'path' => $curpath];
        }
        return [$list, $breadcrumb];
    }

    /**
     * Get listing for a trending files folder using the unified api.
     *
     * @param string $path Folder path.
     * @return array List of $list array and $path array.
     */
    protected function get_listing_trending_unified($path = '') {
        $path = (empty($path)) ? '/' : $path;
        $list = [];
        $unified = $this->get_unified_apiclient();
        $realpath = $path;
        $contents = $unified->get_trending_files($realpath);
        $list = $this->contents_api_response_to_list($contents, $realpath, 'trendingaround', null, false);

        // Generate path.
        $strtrendingfiles = get_string('trendingaround', 'repository_office365');
        $breadcrumb = [['name' => $this->name, 'path' => '/'], ['name' => $strtrendingfiles, 'path' => '/trending/']];
        return [$list, $breadcrumb];
    }

    /**
     * Transform a onedrive API response for a folder into a list parameter that the respository class can understand.
     *
     * @param string $response The response from the API.
     * @param string $path The list path.
     * @param string $clienttype The type of client that the response is from. onedrive/sharepoint.
     * @param string $parentinfo Client type-specific parent information.
     *                               If using the Sharepoint clienttype, this is the parent site URI.
     *                               If using the unifiedgroup clienttype, this is the parent group ID.
     * @param bool $addupload Whether to add the "Upload" file item.
     * @return array A $list array to be used by the respository class in get_listing.
     */
    protected function contents_api_response_to_list($response, $path, $clienttype, $parentinfo = null, $addupload = true) {
        global $OUTPUT, $DB;
        $list = [];
        if ($clienttype === 'onedrive') {
            $pathprefix = '/my'.$path;
            $uploadpathprefix = $pathprefix;
        } else if ($clienttype === 'unified') {
            $pathprefix = '/my';
            $uploadpathprefix = $pathprefix.$path;
        } else if ($clienttype === 'sharepoint') {
            $pathprefix = '/courses'.$path;
            $uploadpathprefix = $pathprefix;
        } else if ($clienttype === 'unifiedgroup') {
            $pathprefix = '/groups'.$path;
            $uploadpathprefix = $pathprefix;
        } else if ($clienttype === 'trendingaround') {
            $pathprefix = '/my';
        } else if ($clienttype === 'office365video') {
            $pathprefix = '/office365video';
            $uploadpathprefix = $pathprefix.$path;
        }

        if ($addupload === true) {
            $list[] = [
                'title' => get_string('upload', 'repository_office365'),
                'path' => $uploadpathprefix.'/upload/',
                'thumbnail' => $OUTPUT->pix_url('a/add_file')->out(false),
                'children' => [],
            ];
        }

        if (isset($response['value'])) {
            foreach ($response['value'] as $content) {
                if ($clienttype === 'unified' || $clienttype === 'unifiedgroup') {
                    $itempath = $pathprefix.'/'.$content['id'];
                    if (isset($content['folder'])) {
                        $list[] = [
                            'title' => $content['name'],
                            'path' => $itempath,
                            'thumbnail' => $OUTPUT->pix_url(file_folder_icon(90))->out(false),
                            'date' => strtotime($content['createdDateTime']),
                            'datemodified' => strtotime($content['lastModifiedDateTime']),
                            'datecreated' => strtotime($content['createdDateTime']),
                            'children' => [],
                        ];
                    } else if (isset($content['file'])) {
                        $url = $content['webUrl'].'?web=1';
                        if ($clienttype === 'unified') {
                            $source = [
                                'id' => $content['id'],
                                'source' => 'onedrive',
                            ];
                        } else if ($clienttype === 'unifiedgroup') {
                            $source = [
                                'id' => $content['id'],
                                'source' => 'onedrivegroup',
                                'groupid' => $parentinfo,
                            ];
                        }

                        $author = '';
                        if (!empty($content['createdBy']['user']['displayName'])) {
                            $author = $content['createdBy']['user']['displayName'];
                            $author = explode(',', $author);
                            $author = $author[0];
                        }

                        $list[] = [
                            'title' => $content['name'],
                            'date' => strtotime($content['createdDateTime']),
                            'datemodified' => strtotime($content['lastModifiedDateTime']),
                            'datecreated' => strtotime($content['createdDateTime']),
                            'size' => $content['size'],
                            'url' => $url,
                            'thumbnail' => $OUTPUT->pix_url(file_extension_icon($content['name'], 90))->out(false),
                            'author' => $author,
                            'source' => $this->pack_reference($source),
                        ];
                    }
                } else if ($clienttype === 'trendingaround') {
                    if (isset($content['folder'])) {
                        $list[] = [
                            'title' => $content['name'],
                            'path' => $itempath = $pathprefix . '/' . $content['name'],
                            'thumbnail' => $OUTPUT->pix_url(file_folder_icon(90))->out(false),
                            'date' => strtotime($content['DateTimeCreated']),
                            'datemodified' => strtotime($content['DateTimeLastModified']),
                            'datecreated' => strtotime($content['DateTimeCreated']),
                            'children' => [],
                        ];
                    } else {
                        $url = $content['webUrl'] . '?web=1';
                        $source = [
                            'id' => $content['@odata.id'],
                            'source' => 'trendingaround',
                        ];

                        $list[] = [
                            'title' => $content['name'],
                            'date' => strtotime($content['DateTimeCreated']),
                            'datemodified' => strtotime($content['DateTimeLastModified']),
                            'datecreated' => strtotime($content['DateTimeCreated']),
                            'url' => $url,
                            'thumbnail' => $OUTPUT->pix_url(file_extension_icon($content['name'], 90))->out(false),
                            'source' => $this->pack_reference($source),
                        ];
                    }
                } else if ($clienttype === 'office365video') {
                    if ($content['odata.type'] === 'SP.Publishing.VideoChannel') {
                        $itempath = $pathprefix.'/'.$content['Id'];
                        $list[] = [
                            'title' => $content['Title'],
                            'path' => $itempath,
                            'thumbnail' => $OUTPUT->pix_url(file_folder_icon(90))->out(false),
                            'children' => [],
                        ];
                    } else if ($content['odata.type'] === 'SP.Publishing.VideoItem') {
                        $itempath = $pathprefix.'/'.$content['ID'];
                        $parseurl = explode('/', $content['Url']);
                        $downloadurl = "https://".$parseurl[2]."/_api/SP.AppContextSite(@target)/Web/"."GetFileByServerRelativeUrl('".$content['ServerRelativeUrl'].
                                "')/"."$"."value?@target='https://".$parseurl[2]."/portals/".$parseurl[4]."'";
                        $url = "https://".$parseurl[2]."/portals/hub/_layouts/15/PointPublishing.aspx?app=video&"."p=p&chid=".$content['ChannelID']."&vid=".$content['ID'];
                        $source = [
                            'id' => $content['odata.id'],
                            'source' => 'office365video',
                            'url' => $url,
                            'downloadurl' => $downloadurl,
                        ];
                        $list[] = [
                            'title' => $content['FileName'],
                            'date' => strtotime($content['CreatedDate']),
                            'datecreated' => strtotime($content['CreatedDate']),
                            'url' => $url,
                            'thumbnail' => $OUTPUT->pix_url(file_extension_icon($content['FileName'], 90))->out(false),
                            'source' => $this->pack_reference($source),
                        ];
                    }
                } else {
                    $itempath = $pathprefix.'/'.$content['name'];
                    if ($content['type'] === 'Folder') {
                        $list[] = [
                            'title' => $content['name'],
                            'path' => $itempath,
                            'thumbnail' => $OUTPUT->pix_url(file_folder_icon(90))->out(false),
                            'date' => strtotime($content['dateTimeCreated']),
                            'datemodified' => strtotime($content['dateTimeLastModified']),
                            'datecreated' => strtotime($content['dateTimeCreated']),
                            'children' => [],
                        ];
                    } else if ($content['type'] === 'File') {
                        $url = $content['webUrl'].'?web=1';
                        $source = [
                            'id' => $content['id'],
                            'source' => ($clienttype === 'sharepoint') ? 'sharepoint' : 'onedrive',
                        ];
                        if ($clienttype === 'sharepoint') {
                            $source['parentsiteuri'] = $parentinfo;
                        }

                        $author = '';
                        if (!empty($content['createdBy']['user']['displayName'])) {
                            $author = $content['createdBy']['user']['displayName'];
                            $author = explode(',', $author);
                            $author = $author[0];
                        }

                        $list[] = [
                            'title' => $content['name'],
                            'date' => strtotime($content['dateTimeCreated']),
                            'datemodified' => strtotime($content['dateTimeLastModified']),
                            'datecreated' => strtotime($content['dateTimeCreated']),
                            'size' => $content['size'],
                            'url' => $url,
                            'thumbnail' => $OUTPUT->pix_url(file_extension_icon($content['name'], 90))->out(false),
                            'author' => $author,
                            'source' => $this->pack_reference($source),
                        ];
                    }
                }
            }
        }
        return $list;
    }

    /**
     * Tells how the file can be picked from this repository
     *
     * Maximum value is FILE_INTERNAL | FILE_EXTERNAL | FILE_REFERENCE
     *
     * @return int
     */
    public function supported_returntypes() {
        return FILE_INTERNAL | FILE_EXTERNAL | FILE_REFERENCE;
    }

    /**
     * Downloads a file from external repository and saves it in temp dir
     *
     * @param string $reference The file reference.
     * @param string $filename filename (without path) to save the downloaded file in the temporary directory, if omitted
     *                         or file already exists the new filename will be generated
     * @return array with elements:
     *   path: internal location of the file
     *   url: URL to the source (from parameters)
     */
    public function get_file($reference, $filename = '') {
        $caller = '\repository_office365::get_file';
        $reference = $this->unpack_reference($reference);

        if ($reference['source'] === 'onedrive') {
            if ($this->unifiedconfigured === true) {
                $sourceclient = $this->get_unified_apiclient();
            } else {
                $sourceclient = $this->get_onedrive_apiclient();
            }
            if (empty($sourceclient)) {
                \local_o365\utils::debug('Could not construct onedrive api.', $caller);
                throw new \moodle_exception('errorwhiledownload', 'repository_office365');
            }
            $file = $sourceclient->get_file_by_id($reference['id']);
        } else if ($reference['source'] === 'onedrivegroup') {
            if ($this->unifiedconfigured === true) {
                $sourceclient = $this->get_unified_apiclient();
            } else {
                \local_o365\utils::debug('Tried to access a onedrive group file while the graph api is disabled.', $caller);
                throw new \moodle_exception('errorwhiledownload', 'repository_office365');
            }
            $file = $sourceclient->get_group_file_by_id($reference['groupid'], $reference['id']);
        } else if ($reference['source'] === 'sharepoint') {
            $sourceclient = $this->get_sharepoint_apiclient();
            if (empty($sourceclient)) {
                \local_o365\utils::debug('Could not construct sharepoint api.', $caller);
                throw new \moodle_exception('errorwhiledownload', 'repository_office365');
            }
            if (isset($reference['parentsiteuri'])) {
                $parentsiteuri = $reference['parentsiteuri'];
            } else {
                $parentsiteuri = $sourceclient->get_moodle_parent_site_uri();
            }
            $sourceclient->set_site($parentsiteuri);
            $file = $sourceclient->get_file_by_id($reference['id']);
        } else if ($reference['source'] === 'trendingaround') {
            if ($this->unifiedconfigured === true) {
                $sourceclient = $this->get_unified_apiclient();
            }
            if (empty($sourceclient)) {
                \local_o365\utils::debug('Could not construct unified api.', $caller);
                throw new \moodle_exception('errorwhiledownload', 'repository_office365');
            }
            $file = $sourceclient->get_file_by_url($reference['url']);
        } else if ($reference['source'] === 'office365video') {
            $sourceclient = $this->get_sharepoint_apiclient();
            if (empty($sourceclient)) {
                \local_o365\utils::debug('Could not construct sharepoint api.', $caller);
                throw new \moodle_exception('errorwhiledownload', 'repository_office365');
            }
            $file = $sourceclient->get_video_file($reference['downloadurl']);
        }

        if (!empty($file)) {
            $path = $this->prepare_file($filename);
            if (!empty($path)) {
                $result = file_put_contents($path, $file);
            }
        }
        if (empty($result)) {
            $errmsg = get_string('errorwhiledownload', 'repository_office365');
            $debugdata = [
                'reference' => $reference,
                'filename' => $filename,
            ];
            \local_o365\utils::debug($errmsg, $caller, $debugdata);
            throw new \moodle_exception('errorwhiledownload', 'repository_office365');
        }
        return ['path' => $path, 'url' => $reference];
    }

    /**
     * Pack file reference information into a string.
     *
     * @param array $reference The information to pack.
     * @return string The packed information.
     */
    protected function pack_reference($reference) {
        return base64_encode(serialize($reference));
    }

    /**
     * Unpack file reference information from a string.
     *
     * @param string $reference The information to unpack.
     * @return array The unpacked information.
     */
    protected function unpack_reference($reference) {
        return unserialize(base64_decode($reference));
    }

    /**
     * Prepare file reference information
     *
     * @param string $source source of the file, returned by repository as 'source' and received back from user (not cleaned)
     * @return string file reference, ready to be stored
     */
    public function get_file_reference($source) {
        $caller = '\repository_office365::get_file_reference';
        $sourceunpacked = $this->unpack_reference($source);
        if (isset($sourceunpacked['source']) && isset($sourceunpacked['id'])) {
            $fileid = $sourceunpacked['id'];
            $filesource = $sourceunpacked['source'];

            $reference = [
                'source' => $filesource,
                'id' => $fileid,
                'url' => '',
            ];

            if (isset($sourceunpacked['url'])) {
                $reference['url'] = $sourceunpacked['url'];
            }
            if (isset($sourceunpacked['downloadurl'])) {
                $reference['downloadurl'] = $sourceunpacked['downloadurl'];
            }

            try {
                if ($filesource === 'onedrive') {
                    if ($this->unifiedconfigured === true) {
                        $sourceclient = $this->get_unified_apiclient();
                        $reference['url'] = $sourceclient->get_sharing_link($fileid);
                    } else {
                        $sourceclient = $this->get_onedrive_apiclient();
                        $filemetadata = $sourceclient->get_file_metadata($fileid);
                        if (isset($filemetadata['webUrl'])) {
                            $reference['url'] = $filemetadata['webUrl'].'?web=1';
                        }
                    }
                } else if ($filesource === 'onedrivegroup') {
                    if ($this->unifiedconfigured !== true) {
                        \local_o365\utils::debug('Tried to access a onedrive group file while the graph api is disabled.', $caller);
                        throw new \moodle_exception('errorwhiledownload', 'repository_office365');
                    }
                    $sourceclient = $this->get_unified_apiclient();
                    $reference['groupid'] = $sourceunpacked['groupid'];
                    $reference['url'] = $sourceclient->get_group_file_sharing_link($sourceunpacked['groupid'], $fileid);
                } else if ($filesource === 'sharepoint') {
                    $sourceclient = $this->get_sharepoint_apiclient();
                    if (isset($sourceunpacked['parentsiteuri'])) {
                        $parentsiteuri = $sourceunpacked['parentsiteuri'];
                    } else {
                        $parentsiteuri = $sourceclient->get_moodle_parent_site_uri();
                    }
                    $sourceclient->set_site($parentsiteuri);
                    $reference['parentsiteuri'] = $parentsiteuri;
                    $filemetadata = $sourceclient->get_file_metadata($fileid);
                    if (isset($filemetadata['webUrl'])) {
                        $reference['url'] = $filemetadata['webUrl'].'?web=1';
                    }
                } else if ($filesource === 'trendingaround') {
                    if ($this->unifiedconfigured !== true) {
                        \local_o365\utils::debug('Tried to access a trending around me file while the graph api is disabled.', $caller);
                        throw new \moodle_exception('errorwhiledownload', 'repository_office365');
                    }
                    $sourceclient = $this->get_unified_apiclient();
                    $filedata = $sourceclient->get_file_data($fileid);
                    if (isset($filedata['@microsoft.graph.downloadUrl'])) {
                        $reference['url'] = $filedata['@microsoft.graph.downloadUrl'];
                    }
                }

            } catch (\Exception $e) {
                $errmsg = 'There was a problem making the API call.';
                $debugdata = [
                    'source' => $filesource,
                    'id' => $fileid,
                    'message' => $e->getMessage(),
                    'e' => $e,
                ];
                \local_o365\utils::debug($errmsg, $caller, $debugdata);
            }

            return $this->pack_reference($reference);
        } else {
            $errmsg = '';
            if (!isset($sourceunpacked['source'])) {
                $errmsg = 'Source is not set.';
            }
            if (isset($sourceunpacked['id'])) {
                $errmsg .= ' id is not set.';
            }
            $debugdata = ['sourceunpacked' => $sourceunpacked];
            \local_o365\utils::debug($errmsg, $caller, $debugdata);
        }
        return $source;
    }

    /**
     * Return file URL, for most plugins, the parameter is the original
     * url, but some plugins use a file id, so we need this function to
     * convert file id to original url.
     *
     * @param string $url the url of file
     * @return string
     */
    public function get_link($url) {
        $reference = $this->unpack_reference($url);
        return $reference['url'];
    }

    /**
     * Determine whether a "send_file" request should be a redirect to the embed URL for a file.
     *
     * @param array $reference The file reference array.
     * @param bool $forcedownload The send_file "forcedownload" param.
     * @return bool True if we should do embedding, false otherwise.
     */
    public function do_embedding($reference, $forcedownload) {
        if ($_SERVER['SCRIPT_NAME'] === '/draftfile.php') {
            return false;
        }
        if (empty($reference['source']) || !in_array($reference['source'], ['onedrive', 'sharepoint'])) {
            return false;
        }
        if (!empty($forcedownload)) {
            return false;
        }
        return true;
    }

    /**
     * Repository method to serve the referenced file
     *
     * @see send_stored_file
     *
     * @param stored_file $storedfile the file that contains the reference
     * @param int $lifetime Number of seconds before the file should expire from caches (null means $CFG->filelifetime)
     * @param int $filter 0 (default)=no filtering, 1=all files, 2=html files only
     * @param bool $forcedownload If true (default false), forces download of file rather than view in browser/plugin
     * @param array $options additional options affecting the file serving
     */
    public function send_file($storedfile, $lifetime = null , $filter = 0, $forcedownload = false, array $options = null) {
        global $USER;
        $caller = '\repository_office365::send_file';
        $reference = $this->unpack_reference($storedfile->get_reference());

        $fileuserid = $storedfile->get_userid();

        if (!isset($reference['source'])) {
            \local_o365\utils::debug('File reference is broken - no source parameter.', 'send_file', $reference);
            send_file_not_found();
            die();
        }

        switch ($reference['source']) {
            case 'sharepoint':
                $sourceclient = $this->get_sharepoint_apiclient(false, $fileuserid);
                if (isset($reference['parentsiteuri'])) {
                    $parentsiteuri = $reference['parentsiteuri'];
                } else {
                    $parentsiteuri = $sourceclient->get_moodle_parent_site_uri();
                }
                $sourceclient->set_site($parentsiteuri);
                if (empty($sourceclient)) {
                    \local_o365\utils::debug('Could not construct api client for user', 'send_file', $fileuserid);
                    send_file_not_found();
                    die();
                }
                $fileinfo = $sourceclient->get_file_metadata($reference['id']);
                break;

            case 'onedrive':
                $sourceclient = $this->get_onedrive_apiclient(false, $fileuserid);
                if (empty($sourceclient)) {
                    \local_o365\utils::debug('Could not construct api client for user', 'send_file', $fileuserid);
                    send_file_not_found();
                    die();
                }
                $fileinfo = $sourceclient->get_file_metadata($reference['id']);
                break;

            case 'onedrivegroup':
                $sourceclient = $this->get_unified_apiclient();
                $fileinfo = $sourceclient->get_group_file_metadata($reference['groupid'], $reference['id']);
                break;

            case 'office365video':
                break;

            default:
                \local_o365\utils::debug('File reference is broken - invalid source parameter.', 'send_file', $reference);
                send_file_not_found();
                die();
        }

        // Do embedding if relevant.
        $doembed = $this->do_embedding($reference, $forcedownload);
        if ($doembed === true) {
            if (\local_o365\utils::is_o365_connected($USER->id) !== true) {
                // Embedding currently only supported for logged-in Office 365 users.
                echo get_string('erroro365required', 'repository_office365');
                die();
            }
            if (!empty($sourceclient)) {
                if (isset($fileinfo['webUrl'])) {
                    $fileurl = $fileinfo['webUrl'];
                } else {
                    $fileurl = (isset($reference['url'])) ? $reference['url'] : '';
                }

                if (empty($fileurl)) {
                    $errstr = 'Embed was requested, but could not get file info to complete request.';
                    \local_o365\utils::debug($errstr, 'send_file', ['reference' => $reference, 'fileinfo' => $fileinfo]);
                } else {
                    try {
                        $embedurl = $sourceclient->get_embed_url($reference['id'], $fileurl);
                        $embedurl = (isset($embedurl['value'])) ? $embedurl['value'] : '';
                    } catch (\Exception $e) {
                        // Note: exceptions will already be logged in get_embed_url.
                        $embedurl = '';
                    }
                    if (!empty($embedurl)) {
                        redirect($embedurl);
                    } else if (!empty($fileurl)) {
                        redirect($fileurl);
                    } else {
                        $errstr = 'Embed was requested, but could not complete.';
                        \local_o365\utils::debug($errstr, 'send_file', $reference);
                    }
                }
            } else {
                \local_o365\utils::debug('Could not construct OneDrive client for system api user.', 'send_file');
            }
        }

        if ($reference['source'] === 'office365video') {
            redirect($reference['url']);
        }

        redirect($fileinfo['webUrl']);
    }

    /**
     * Validate Admin Settings Moodle form
     *
     * @static
     * @param moodleform $mform Moodle form (passed by reference)
     * @param array $data array of ("fieldname"=>value) of submitted data
     * @param array $errors array of ("fieldname"=>errormessage) of errors
     * @return array array of errors
     */
    public static function type_form_validation($mform, $data, $errors) {
        global $CFG;
        if (\local_o365\utils::is_configured() !== true) {
            array_push($errors, get_string('notconfigured', 'repository_office365', $CFG->wwwroot));
        }
        return $errors;
    }

    /**
     * Setup repistory form.
     *
     * @param moodleform $mform Moodle form (passed by reference)
     * @param string $classname repository class name
     */
    public static function type_config_form($mform, $classname = 'repository') {
        global $CFG;
        $a = new stdClass;
        if (\local_o365\utils::is_configured() !== true) {
            $mform->addElement('static', null, '', get_string('notconfigured', 'repository_office365', $CFG->wwwroot));
        }
        parent::type_config_form($mform);
    }
}
