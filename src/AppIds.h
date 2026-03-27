#pragma once

// Window and control identifiers.
constexpr int ID_APP_ICON = 101;

constexpr int IDC_SIDEBAR = 1001;
constexpr int IDC_SETTINGS_BUTTON = 1002;
constexpr int IDC_CONTENT_TITLE = 1003;
constexpr int IDC_NAV_CUT = 1004;
constexpr int IDC_NAV_CONCAT = 1005;
constexpr int IDC_NAV_FADE = 1006;
constexpr int IDC_NAV_CONVERT = 1007;

constexpr int IDC_CUT_GROUP = 1100;
constexpr int IDC_CUT_INPUT_LABEL = 1101;
constexpr int IDC_CUT_INPUT_EDIT = 1102;
constexpr int IDC_CUT_INPUT_BROWSE = 1103;
constexpr int IDC_CUT_START_LABEL = 1104;
constexpr int IDC_CUT_START_TRACK = 1105;
constexpr int IDC_CUT_END_LABEL = 1106;
constexpr int IDC_CUT_END_TRACK = 1107;
constexpr int IDC_CUT_RUN_BUTTON = 1111;
constexpr int IDC_CUT_LOG_EDIT = 1112;
constexpr int IDC_CONCAT_LIST = 1113;
constexpr int IDC_CONCAT_CLEAR_BUTTON = 1124;
constexpr int IDC_FADE_IN_TRACK = 1125;
constexpr int IDC_FADE_OUT_TRACK = 1126;
constexpr int IDC_CONVERT_MODE_VIDEO = 1127;
constexpr int IDC_CONVERT_MODE_AUDIO = 1128;
constexpr int IDC_CONVERT_HINT = 1129;
constexpr int IDC_CONVERT_INFO_LABEL_FILE = 1130;
constexpr int IDC_CONVERT_INFO_VALUE_FILE = 1131;
constexpr int IDC_CONVERT_INFO_LABEL_FORMAT = 1132;
constexpr int IDC_CONVERT_INFO_VALUE_FORMAT = 1133;
constexpr int IDC_CONVERT_INFO_LABEL_VCODEC = 1134;
constexpr int IDC_CONVERT_INFO_VALUE_VCODEC = 1135;
constexpr int IDC_CONVERT_INFO_LABEL_VBITRATE = 1136;
constexpr int IDC_CONVERT_INFO_VALUE_VBITRATE = 1137;
constexpr int IDC_CONVERT_INFO_LABEL_RESOLUTION = 1138;
constexpr int IDC_CONVERT_INFO_VALUE_RESOLUTION = 1139;
constexpr int IDC_CONVERT_INFO_LABEL_ACODEC = 1140;
constexpr int IDC_CONVERT_INFO_VALUE_ACODEC = 1141;
constexpr int IDC_CONVERT_INFO_LABEL_ABITRATE = 1142;
constexpr int IDC_CONVERT_INFO_VALUE_ABITRATE = 1143;
constexpr int IDC_CONVERT_TO_MP3_BUTTON = 1144;
constexpr int IDC_CONVERT_TO_MP4_BUTTON = 1145;
constexpr int IDC_CUT_START_DECREASE = 1114;
constexpr int IDC_CUT_START_VALUE = 1115;
constexpr int IDC_CUT_START_INCREASE = 1116;
constexpr int IDC_CUT_END_DECREASE = 1117;
constexpr int IDC_CUT_END_VALUE = 1118;
constexpr int IDC_CUT_END_INCREASE = 1119;
constexpr int IDC_CUT_OPEN_FOLDER_BUTTON = 1120;
constexpr int IDC_CUT_PREVIEW_BUTTON = 1121;
constexpr int IDC_CUT_VIDEO_PREVIEW = 1122;
constexpr int IDC_CUT_LOG_TOGGLE = 1123;

constexpr int IDC_SETTINGS_DIALOG = 1200;
constexpr int IDC_SETTINGS_PATH_LABEL = 1201;
constexpr int IDC_SETTINGS_PATH_EDIT = 1202;
constexpr int IDC_SETTINGS_BROWSE = 1203;
constexpr int IDC_SETTINGS_OK = 1204;
constexpr int IDC_SETTINGS_CANCEL = 1205;
constexpr int IDC_SETTINGS_PROBE_LABEL = 1206;
constexpr int IDC_SETTINGS_PROBE_EDIT = 1207;
constexpr int IDC_SETTINGS_PROBE_BROWSE = 1208;
constexpr int IDC_SETTINGS_SUFFIX_LABEL = 1209;
constexpr int IDC_SETTINGS_SUFFIX_EDIT = 1210;
constexpr int IDC_SETTINGS_SAVE_SOURCE_CHECK = 1211;

constexpr int WM_APP_RUN_LOG = WM_APP + 1;
constexpr int WM_APP_RUN_FINISHED = WM_APP + 2;
constexpr int WM_APP_DURATION_PROBED = WM_APP + 3;
constexpr int WM_APP_FADE_ITEM_STATUS = WM_APP + 4;
