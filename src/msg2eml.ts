import * as iconvLite from 'iconv-lite';
import { deEncapsulateSync } from 'rtf-stream-parser';
var CFB = require('cfb');
var eml_format = require('eml-format');
var moment = require('moment');
const { decompressRTF } = require('@kenjiuno/decompressrtf');
var bigInt = require("big-integer");

var property_tags = new Array(0x3F08 + 1);
property_tags[0x01] = ['ACKNOWLEDGEMENT_MODE', 'I4'];
property_tags[0x02] = ['ALTERNATE_RECIPIENT_ALLOWED', 'BOOLEAN'];
property_tags[0x03] = ['AUTHORIZING_USERS', 'BINARY'];
// Comment on an automatically forwarded message
property_tags[0x04] = ['AUTO_FORWARD_COMMENT', 'STRING'];
// Whether a message has been automatically forwarded
property_tags[0x05] = ['AUTO_FORWARDED', 'BOOLEAN'];
property_tags[0x06] = ['CONTENT_CONFIDENTIALITY_ALGORITHM_ID', 'BINARY'];
property_tags[0x07] = ['CONTENT_CORRELATOR', 'BINARY'];
property_tags[0x08] = ['CONTENT_IDENTIFIER', 'STRING'];
// MIME content length
property_tags[0x09] = ['CONTENT_LENGTH', 'I4'];
property_tags[0x0A] = ['CONTENT_RETURN_REQUESTED', 'BOOLEAN'];
property_tags[0x0B] = ['CONVERSATION_KEY', 'BINARY'];
property_tags[0x0C] = ['CONVERSION_EITS', 'BINARY'];
property_tags[0x0D] = ['CONVERSION_WITH_LOSS_PROHIBITED', 'BOOLEAN'];
property_tags[0x0E] = ['CONVERTED_EITS', 'BINARY'];
// Time to deliver for delayed delivery messages
property_tags[0x0F] = ['DEFERRED_DELIVERY_TIME', 'SYSTIME'];
property_tags[0x10] = ['DELIVER_TIME', 'SYSTIME'];
// Reason a message was discarded
property_tags[0x11] = ['DISCARD_REASON', 'I4'];
property_tags[0x12] = ['DISCLOSURE_OF_RECIPIENTS', 'BOOLEAN'];
property_tags[0x13] = ['DL_EXPANSION_HISTORY', 'BINARY'];
property_tags[0x14] = ['DL_EXPANSION_PROHIBITED', 'BOOLEAN'];
property_tags[0x15] = ['EXPIRY_TIME', 'SYSTIME'];
property_tags[0x16] = ['IMPLICIT_CONVERSION_PROHIBITED', 'BOOLEAN'];
// Message importance
property_tags[0x17] = ['IMPORTANCE', 'I4'];
property_tags[0x18] = ['IPM_ID', 'BINARY'];
property_tags[0x19] = ['LATEST_DELIVERY_TIME', 'SYSTIME'];
property_tags[0x1A] = ['MESSAGE_CLASS', 'STRING'];
property_tags[0x1B] = ['MESSAGE_DELIVERY_ID', 'BINARY'];
property_tags[0x1E] = ['MESSAGE_SECURITY_LABEL', 'BINARY'];
property_tags[0x1F] = ['OBSOLETED_IPMS', 'BINARY'];
// Person a message was originally for
property_tags[0x20] = ['ORIGINALLY_INTENDED_RECIPIENT_NAME', 'BINARY'];
property_tags[0x21] = ['ORIGINAL_EITS', 'BINARY'];
property_tags[0x22] = ['ORIGINATOR_CERTIFICATE', 'BINARY'];
property_tags[0x23] = ['ORIGINATOR_DELIVERY_REPORT_REQUESTED', 'BOOLEAN'];
// Address of the message sender
property_tags[0x24] = ['ORIGINATOR_RETURN_ADDRESS', 'BINARY'];
property_tags[0x25] = ['PARENT_KEY', 'BINARY'];
property_tags[0x26] = ['PRIORITY', 'I4'];
property_tags[0x27] = ['ORIGIN_CHECK', 'BINARY'];
property_tags[0x28] = ['PROOF_OF_SUBMISSION_REQUESTED', 'BOOLEAN'];
// Whether a read receipt is desired
property_tags[0x29] = ['READ_RECEIPT_REQUESTED', 'BOOLEAN'];
// Time a message was received
property_tags[0x2A] = ['RECEIPT_TIME', 'SYSTIME'];
property_tags[0x2B] = ['RECIPIENT_REASSIGNMENT_PROHIBITED', 'BOOLEAN'];
property_tags[0x2C] = ['REDIRECTION_HISTORY', 'BINARY'];
property_tags[0x2D] = ['RELATED_IPMS', 'BINARY'];
// Sensitivity of the original message
property_tags[0x2E] = ['ORIGINAL_SENSITIVITY', 'I4'];
property_tags[0x2F] = ['LANGUAGES', 'STRING'];
property_tags[0x30] = ['REPLY_TIME', 'SYSTIME'];
property_tags[0x31] = ['REPORT_TAG', 'BINARY'];
property_tags[0x32] = ['REPORT_TIME', 'SYSTIME'];
property_tags[0x33] = ['RETURNED_IPM', 'BOOLEAN'];
property_tags[0x34] = ['SECURITY', 'I4'];
property_tags[0x35] = ['INCOMPLETE_COPY', 'BOOLEAN'];
property_tags[0x36] = ['SENSITIVITY', 'I4'];
// The message subject
property_tags[0x37] = ['SUBJECT', 'STRING'];
property_tags[0x38] = ['SUBJECT_IPM', 'BINARY'];
property_tags[0x39] = ['CLIENT_SUBMIT_TIME', 'SYSTIME'];
property_tags[0x3A] = ['REPORT_NAME', 'STRING'];
property_tags[0x3B] = ['SENT_REPRESENTING_SEARCH_KEY', 'BINARY'];
property_tags[0x3C] = ['X400_CONTENT_TYPE', 'BINARY'];
property_tags[0x3D] = ['SUBJECT_PREFIX', 'STRING'];
property_tags[0x3E] = ['NON_RECEIPT_REASON', 'I4'];
property_tags[0x3F] = ['RECEIVED_BY_ENTRYID', 'BINARY'];
// Received by: entry
property_tags[0x40] = ['RECEIVED_BY_NAME', 'STRING'];
property_tags[0x41] = ['SENT_REPRESENTING_ENTRYID', 'BINARY'];
property_tags[0x42] = ['SENT_REPRESENTING_NAME', 'STRING'];
property_tags[0x43] = ['RCVD_REPRESENTING_ENTRYID', 'BINARY'];
property_tags[0x44] = ['RCVD_REPRESENTING_NAME', 'STRING'];
property_tags[0x45] = ['REPORT_ENTRYID', 'BINARY'];
property_tags[0x46] = ['READ_RECEIPT_ENTRYID', 'BINARY'];
property_tags[0x47] = ['MESSAGE_SUBMISSION_ID', 'BINARY'];
property_tags[0x48] = ['PROVIDER_SUBMIT_TIME', 'SYSTIME'];
// Subject of the original message
property_tags[0x49] = ['ORIGINAL_SUBJECT', 'STRING'];
property_tags[0x4A] = ['DISC_VAL', 'BOOLEAN'];
property_tags[0x4B] = ['ORIG_MESSAGE_CLASS', 'STRING'];
property_tags[0x4C] = ['ORIGINAL_AUTHOR_ENTRYID', 'BINARY'];
// Author of the original message
property_tags[0x4D] = ['ORIGINAL_AUTHOR_NAME', 'STRING'];
// Time the original message was submitted
property_tags[0x4E] = ['ORIGINAL_SUBMIT_TIME', 'SYSTIME'];
property_tags[0x4F] = ['REPLY_RECIPIENT_ENTRIES', 'BINARY'];
property_tags[0x50] = ['REPLY_RECIPIENT_NAMES', 'STRING'];
property_tags[0x51] = ['RECEIVED_BY_SEARCH_KEY', 'BINARY'];
property_tags[0x52] = ['RCVD_REPRESENTING_SEARCH_KEY', 'BINARY'];
property_tags[0x53] = ['READ_RECEIPT_SEARCH_KEY', 'BINARY'];
property_tags[0x54] = ['REPORT_SEARCH_KEY', 'BINARY'];
property_tags[0x55] = ['ORIGINAL_DELIVERY_TIME', 'SYSTIME'];
property_tags[0x56] = ['ORIGINAL_AUTHOR_SEARCH_KEY', 'BINARY'];
property_tags[0x57] = ['MESSAGE_TO_ME', 'BOOLEAN'];
property_tags[0x58] = ['MESSAGE_CC_ME', 'BOOLEAN'];
property_tags[0x59] = ['MESSAGE_RECIP_ME', 'BOOLEAN'];
// Sender of the original message
property_tags[0x5A] = ['ORIGINAL_SENDER_NAME', 'STRING'];
property_tags[0x5B] = ['ORIGINAL_SENDER_ENTRYID', 'BINARY'];
property_tags[0x5C] = ['ORIGINAL_SENDER_SEARCH_KEY', 'BINARY'];
property_tags[0x5D] = ['ORIGINAL_SENT_REPRESENTING_NAME', 'STRING'];
property_tags[0x5E] = ['ORIGINAL_SENT_REPRESENTING_ENTRYID', 'BINARY'];
property_tags[0x5F] = ['ORIGINAL_SENT_REPRESENTING_SEARCH_KEY', 'BINARY'];
property_tags[0x60] = ['START_DATE', 'SYSTIME'];
property_tags[0x61] = ['END_DATE', 'SYSTIME'];
property_tags[0x62] = ['OWNER_APPT_ID', 'I4'];
// Whether a response to the message is desired
property_tags[0x63] = ['RESPONSE_REQUESTED', 'BOOLEAN'];
property_tags[0x64] = ['SENT_REPRESENTING_ADDRTYPE', 'STRING'];
property_tags[0x65] = ['SENT_REPRESENTING_EMAIL_ADDRESS', 'STRING'];
property_tags[0x66] = ['ORIGINAL_SENDER_ADDRTYPE', 'STRING'];
// Email of the original message sender
property_tags[0x67] = ['ORIGINAL_SENDER_EMAIL_ADDRESS', 'STRING'];
property_tags[0x68] = ['ORIGINAL_SENT_REPRESENTING_ADDRTYPE', 'STRING'];
property_tags[0x69] = ['ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS', 'STRING'];
property_tags[0x70] = ['CONVERSATION_TOPIC', 'STRING'];
property_tags[0x71] = ['CONVERSATION_INDEX', 'BINARY'];
property_tags[0x72] = ['ORIGINAL_DISPLAY_BCC', 'STRING'];
property_tags[0x73] = ['ORIGINAL_DISPLAY_CC', 'STRING'];
property_tags[0x74] = ['ORIGINAL_DISPLAY_TO', 'STRING'];
property_tags[0x75] = ['RECEIVED_BY_ADDRTYPE', 'STRING'];
property_tags[0x76] = ['RECEIVED_BY_EMAIL_ADDRESS', 'STRING'];
property_tags[0x77] = ['RCVD_REPRESENTING_ADDRTYPE', 'STRING'];
property_tags[0x78] = ['RCVD_REPRESENTING_EMAIL_ADDRESS', 'STRING'];
property_tags[0x79] = ['ORIGINAL_AUTHOR_ADDRTYPE', 'STRING'];
property_tags[0x7A] = ['ORIGINAL_AUTHOR_EMAIL_ADDRESS', 'STRING'];
property_tags[0x7B] = ['ORIGINALLY_INTENDED_RECIP_ADDRTYPE', 'STRING'];
property_tags[0x7C] = ['ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS', 'STRING'];
property_tags[0x7D] = ['TRANSPORT_MESSAGE_HEADERS', 'STRING'];
property_tags[0x7E] = ['DELEGATION', 'BINARY'];
property_tags[0x7F] = ['TNEF_CORRELATION_KEY', 'BINARY'];
property_tags[0x1000] = ['BODY', 'STRING'];
property_tags[0x1001] = ['REPORT_TEXT', 'STRING'];
property_tags[0x1002] = ['ORIGINATOR_AND_DL_EXPANSION_HISTORY', 'BINARY'];
property_tags[0x1003] = ['REPORTING_DL_NAME', 'BINARY'];
property_tags[0x1004] = ['REPORTING_MTA_CERTIFICATE', 'BINARY'];
property_tags[0x1006] = ['RTF_SYNC_BODY_CRC', 'I4'];
property_tags[0x1007] = ['RTF_SYNC_BODY_COUNT', 'I4'];
property_tags[0x1008] = ['RTF_SYNC_BODY_TAG', 'STRING'];
property_tags[0x1009] = ['RTF_COMPRESSED', 'BINARY'];
property_tags[0x1010] = ['RTF_SYNC_PREFIX_COUNT', 'I4'];
property_tags[0x1011] = ['RTF_SYNC_TRAILING_COUNT', 'I4'];
property_tags[0x1012] = ['ORIGINALLY_INTENDED_RECIP_ENTRYID', 'BINARY'];
property_tags[0x0C00] = ['CONTENT_INTEGRITY_CHECK', 'BINARY'];
property_tags[0x0C01] = ['EXPLICIT_CONVERSION', 'I4'];
property_tags[0x0C02] = ['IPM_RETURN_REQUESTED', 'BOOLEAN'];
property_tags[0x0C03] = ['MESSAGE_TOKEN', 'BINARY'];
property_tags[0x0C04] = ['NDR_REASON_CODE', 'I4'];
property_tags[0x0C05] = ['NDR_DIAG_CODE', 'I4'];
property_tags[0x0C06] = ['NON_RECEIPT_NOTIFICATION_REQUESTED', 'BOOLEAN'];
property_tags[0x0C07] = ['DELIVERY_POINT', 'I4'];
property_tags[0x0C08] = ['ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED', 'BOOLEAN'];
property_tags[0x0C09] = ['ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT', 'BINARY'];
property_tags[0x0C0A] = ['PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY', 'BOOLEAN'];
property_tags[0x0C0B] = ['PHYSICAL_DELIVERY_MODE', 'I4'];
property_tags[0x0C0C] = ['PHYSICAL_DELIVERY_REPORT_REQUEST', 'I4'];
property_tags[0x0C0D] = ['PHYSICAL_FORWARDING_ADDRESS', 'BINARY'];
property_tags[0x0C0E] = ['PHYSICAL_FORWARDING_ADDRESS_REQUESTED', 'BOOLEAN'];
property_tags[0x0C0F] = ['PHYSICAL_FORWARDING_PROHIBITED', 'BOOLEAN'];
property_tags[0x0C10] = ['PHYSICAL_RENDITION_ATTRIBUTES', 'BINARY'];
property_tags[0x0C11] = ['PROOF_OF_DELIVERY', 'BINARY'];
property_tags[0x0C12] = ['PROOF_OF_DELIVERY_REQUESTED', 'BOOLEAN'];
property_tags[0x0C13] = ['RECIPIENT_CERTIFICATE', 'BINARY'];
property_tags[0x0C14] = ['RECIPIENT_NUMBER_FOR_ADVICE', 'STRING'];
property_tags[0x0C15] = ['RECIPIENT_TYPE', 'I4'];
property_tags[0x0C16] = ['REGISTERED_MAIL_TYPE', 'I4'];
property_tags[0x0C17] = ['REPLY_REQUESTED', 'BOOLEAN'];
property_tags[0x0C18] = ['REQUESTED_DELIVERY_METHOD', 'I4'];
property_tags[0x0C19] = ['SENDER_ENTRYID', 'BINARY'];
property_tags[0x0C1A] = ['SENDER_NAME', 'STRING'];
property_tags[0x0C1B] = ['SUPPLEMENTARY_INFO', 'STRING'];
property_tags[0x0C1C] = ['TYPE_OF_MTS_USER', 'I4'];
property_tags[0x0C1D] = ['SENDER_SEARCH_KEY', 'BINARY'];
property_tags[0x0C1E] = ['SENDER_ADDRTYPE', 'STRING'];
property_tags[0x0C1F] = ['SENDER_EMAIL_ADDRESS', 'STRING'];
property_tags[0x0E00] = ['CURRENT_VERSION', 'I8'];
property_tags[0x0E01] = ['DELETE_AFTER_SUBMIT', 'BOOLEAN'];
property_tags[0x0E02] = ['DISPLAY_BCC', 'STRING'];
property_tags[0x0E03] = ['DISPLAY_CC', 'STRING'];
property_tags[0x0E04] = ['DISPLAY_TO', 'STRING'];
property_tags[0x0E05] = ['PARENT_DISPLAY', 'STRING'];
property_tags[0x0E06] = ['MESSAGE_DELIVERY_TIME', 'SYSTIME'];
property_tags[0x0E07] = ['MESSAGE_FLAGS', 'I4'];
property_tags[0x0E08] = ['MESSAGE_SIZE', 'I4'];
property_tags[0x0E09] = ['PARENT_ENTRYID', 'BINARY'];
property_tags[0x0E0A] = ['SENTMAIL_ENTRYID', 'BINARY'];
property_tags[0x0E0C] = ['CORRELATE', 'BOOLEAN'];
property_tags[0x0E0D] = ['CORRELATE_MTSID', 'BINARY'];
property_tags[0x0E0E] = ['DISCRETE_VALUES', 'BOOLEAN'];
property_tags[0x0E0F] = ['RESPONSIBILITY', 'BOOLEAN'];
property_tags[0x0E10] = ['SPOOLER_STATUS', 'I4'];
property_tags[0x0E11] = ['TRANSPORT_STATUS', 'I4'];
property_tags[0x0E12] = ['MESSAGE_RECIPIENTS', 'OBJECT'];
property_tags[0x0E13] = ['MESSAGE_ATTACHMENTS', 'OBJECT'];
property_tags[0x0E14] = ['SUBMIT_FLAGS', 'I4'];
property_tags[0x0E15] = ['RECIPIENT_STATUS', 'I4'];
property_tags[0x0E16] = ['TRANSPORT_KEY', 'I4'];
property_tags[0x0E17] = ['MSG_STATUS', 'I4'];
property_tags[0x0E18] = ['MESSAGE_DOWNLOAD_TIME', 'I4'];
property_tags[0x0E19] = ['CREATION_VERSION', 'I8'];
property_tags[0x0E1A] = ['MODIFY_VERSION', 'I8'];
property_tags[0x0E1B] = ['HASATTACH', 'BOOLEAN'];
property_tags[0x0E1D] = ['NORMALIZED_SUBJECT', 'STRING'];
property_tags[0x0E1F] = ['RTF_IN_SYNC', 'BOOLEAN'];
property_tags[0x0E20] = ['ATTACH_SIZE', 'I4'];
property_tags[0x0E21] = ['ATTACH_NUM', 'I4'];
property_tags[0x0E22] = ['PREPROCESS', 'BOOLEAN'];
property_tags[0x0E25] = ['ORIGINATING_MTA_CERTIFICATE', 'BINARY'];
property_tags[0x0E26] = ['PROOF_OF_SUBMISSION', 'BINARY'];
// A unique identifier for editing the properties of a MAPI object
property_tags[0x0FFF] = ['ENTRYID', 'BINARY'];
// The type of an object
property_tags[0x0FFE] = ['OBJECT_TYPE', 'I4'];
property_tags[0x0FFD] = ['ICON', 'BINARY'];
property_tags[0x0FFC] = ['MINI_ICON', 'BINARY'];
property_tags[0x0FFB] = ['STORE_ENTRYID', 'BINARY'];
property_tags[0x0FFA] = ['STORE_RECORD_KEY', 'BINARY'];
// Binary identifer for an individual object
property_tags[0x0FF9] = ['RECORD_KEY', 'BINARY'];
property_tags[0x0FF8] = ['MAPPING_SIGNATURE', 'BINARY'];
property_tags[0x0FF7] = ['ACCESS_LEVEL', 'I4'];
// The primary key of a column in a table
property_tags[0x0FF6] = ['INSTANCE_KEY', 'BINARY'];
property_tags[0x0FF5] = ['ROW_TYPE', 'I4'];
property_tags[0x0FF4] = ['ACCESS', 'I4'];
property_tags[0x3000] = ['ROWID', 'I4'];
// The name to display for a given MAPI object
property_tags[0x3001] = ['DISPLAY_NAME', 'STRING'];
property_tags[0x3002] = ['ADDRTYPE', 'STRING'];
// An email address
property_tags[0x3003] = ['EMAIL_ADDRESS', 'STRING'];
// A comment field
property_tags[0x3004] = ['COMMENT', 'STRING'];
property_tags[0x3005] = ['DEPTH', 'I4'];
// Provider-defined display name for a service provider
property_tags[0x3006] = ['PROVIDER_DISPLAY', 'STRING'];
// The time an object was created
property_tags[0x3007] = ['CREATION_TIME', 'SYSTIME'];
// The time an object was last modified
property_tags[0x3008] = ['LAST_MODIFICATION_TIME', 'SYSTIME'];
// Flags describing a service provider, message service, or status object
property_tags[0x3009] = ['RESOURCE_FLAGS', 'I4'];
// The name of a provider dll, minus any "32" suffix and ".dll"
property_tags[0x300A] = ['PROVIDER_DLL_NAME', 'STRING'];
property_tags[0x300B] = ['SEARCH_KEY', 'BINARY'];
property_tags[0x300C] = ['PROVIDER_UID', 'BINARY'];
property_tags[0x300D] = ['PROVIDER_ORDINAL', 'I4'];
property_tags[0x3301] = ['FORM_VERSION', 'STRING'];
property_tags[0x3302] = ['FORM_CLSID', 'CLSID'];
property_tags[0x3303] = ['FORM_CONTACT_NAME', 'STRING'];
property_tags[0x3304] = ['FORM_CATEGORY', 'STRING'];
property_tags[0x3305] = ['FORM_CATEGORY_SUB', 'STRING'];
property_tags[0x3306] = ['FORM_HOST_MAP', 'MV_LONG'];
property_tags[0x3307] = ['FORM_HIDDEN', 'BOOLEAN'];
property_tags[0x3308] = ['FORM_DESIGNER_NAME', 'STRING'];
property_tags[0x3309] = ['FORM_DESIGNER_GUID', 'CLSID'];
property_tags[0x330A] = ['FORM_MESSAGE_BEHAVIOR', 'I4'];
// Is this row the default message store?
property_tags[0x3400] = ['DEFAULT_STORE', 'BOOLEAN'];
property_tags[0x340D] = ['STORE_SUPPORT_MASK', 'I4'];
property_tags[0x340E] = ['STORE_STATE', 'I4'];
property_tags[0x3410] = ['IPM_SUBTREE_SEARCH_KEY', 'BINARY'];
property_tags[0x3411] = ['IPM_OUTBOX_SEARCH_KEY', 'BINARY'];
property_tags[0x3412] = ['IPM_WASTEBASKET_SEARCH_KEY', 'BINARY'];
property_tags[0x3413] = ['IPM_SENTMAIL_SEARCH_KEY', 'BINARY'];
// Provder-defined message store type
property_tags[0x3414] = ['MDB_PROVIDER', 'BINARY'];
property_tags[0x3415] = ['RECEIVE_FOLDER_SETTINGS', 'OBJECT'];
property_tags[0x35DF] = ['VALID_FOLDER_MASK', 'I4'];
property_tags[0x35E0] = ['IPM_SUBTREE_ENTRYID', 'BINARY'];
property_tags[0x35E2] = ['IPM_OUTBOX_ENTRYID', 'BINARY'];
property_tags[0x35E3] = ['IPM_WASTEBASKET_ENTRYID', 'BINARY'];
property_tags[0x35E4] = ['IPM_SENTMAIL_ENTRYID', 'BINARY'];
property_tags[0x35E5] = ['VIEWS_ENTRYID', 'BINARY'];
property_tags[0x35E6] = ['COMMON_VIEWS_ENTRYID', 'BINARY'];
property_tags[0x35E7] = ['FINDER_ENTRYID', 'BINARY'];
property_tags[0x3600] = ['CONTAINER_FLAGS', 'I4'];
property_tags[0x3601] = ['FOLDER_TYPE', 'I4'];
property_tags[0x3602] = ['CONTENT_COUNT', 'I4'];
property_tags[0x3603] = ['CONTENT_UNREAD', 'I4'];
property_tags[0x3604] = ['CREATE_TEMPLATES', 'OBJECT'];
property_tags[0x3605] = ['DETAILS_TABLE', 'OBJECT'];
property_tags[0x3607] = ['SEARCH', 'OBJECT'];
property_tags[0x3609] = ['SELECTABLE', 'BOOLEAN'];
property_tags[0x360A] = ['SUBFOLDERS', 'BOOLEAN'];
property_tags[0x360B] = ['STATUS', 'I4'];
property_tags[0x360C] = ['ANR', 'STRING'];
property_tags[0x360D] = ['CONTENTS_SORT_ORDER', 'MV_LONG'];
property_tags[0x360E] = ['CONTAINER_HIERARCHY', 'OBJECT'];
property_tags[0x360F] = ['CONTAINER_CONTENTS', 'OBJECT'];
property_tags[0x3610] = ['FOLDER_ASSOCIATED_CONTENTS', 'OBJECT'];
property_tags[0x3611] = ['DEF_CREATE_DL', 'BINARY'];
property_tags[0x3612] = ['DEF_CREATE_MAILUSER', 'BINARY'];
property_tags[0x3613] = ['CONTAINER_CLASS', 'STRING'];
property_tags[0x3614] = ['CONTAINER_MODIFY_VERSION', 'I8'];
property_tags[0x3615] = ['AB_PROVIDER_ID', 'BINARY'];
property_tags[0x3616] = ['DEFAULT_VIEW_ENTRYID', 'BINARY'];
property_tags[0x3617] = ['ASSOC_CONTENT_COUNT', 'I4'];
property_tags[0x3700] = ['ATTACHMENT_X400_PARAMETERS', 'BINARY'];
property_tags[0x3701] = ['ATTACH_DATA_OBJ', 'OBJECT'];
property_tags[0x3701] = ['ATTACH_DATA_BIN', 'BINARY'];
property_tags[0x3702] = ['ATTACH_ENCODING', 'BINARY'];
property_tags[0x3703] = ['ATTACH_EXTENSION', 'STRING'];
property_tags[0x3704] = ['ATTACH_FILENAME', 'STRING'];
property_tags[0x3705] = ['ATTACH_METHOD', 'I4'];
property_tags[0x3707] = ['ATTACH_LONG_FILENAME', 'STRING'];
property_tags[0x3708] = ['ATTACH_PATHNAME', 'STRING'];
property_tags[0x370A] = ['ATTACH_TAG', 'BINARY'];
property_tags[0x370B] = ['RENDERING_POSITION', 'I4'];
property_tags[0x370C] = ['ATTACH_TRANSPORT_NAME', 'STRING'];
property_tags[0x370D] = ['ATTACH_LONG_PATHNAME', 'STRING'];
property_tags[0x370E] = ['ATTACH_MIME_TAG', 'STRING'];
property_tags[0x370F] = ['ATTACH_ADDITIONAL_INFO', 'BINARY'];
property_tags[0x3900] = ['DISPLAY_TYPE', 'I4'];
property_tags[0x3902] = ['TEMPLATEID', 'BINARY'];
property_tags[0x3904] = ['PRIMARY_CAPABILITY', 'BINARY'];
property_tags[0x39FF] = ['7BIT_DISPLAY_NAME', 'STRING'];
property_tags[0x3A00] = ['ACCOUNT', 'STRING'];
property_tags[0x3A01] = ['ALTERNATE_RECIPIENT', 'BINARY'];
property_tags[0x3A02] = ['CALLBACK_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A03] = ['CONVERSION_PROHIBITED', 'BOOLEAN'];
property_tags[0x3A04] = ['DISCLOSE_RECIPIENTS', 'BOOLEAN'];
property_tags[0x3A05] = ['GENERATION', 'STRING'];
property_tags[0x3A06] = ['GIVEN_NAME', 'STRING'];
property_tags[0x3A07] = ['GOVERNMENT_ID_NUMBER', 'STRING'];
property_tags[0x3A08] = ['BUSINESS_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A09] = ['HOME_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A0A] = ['INITIALS', 'STRING'];
property_tags[0x3A0B] = ['KEYWORD', 'STRING'];
property_tags[0x3A0C] = ['LANGUAGE', 'STRING'];
property_tags[0x3A0D] = ['LOCATION', 'STRING'];
property_tags[0x3A0E] = ['MAIL_PERMISSION', 'BOOLEAN'];
property_tags[0x3A0F] = ['MHS_COMMON_NAME', 'STRING'];
property_tags[0x3A10] = ['ORGANIZATIONAL_ID_NUMBER', 'STRING'];
property_tags[0x3A11] = ['SURNAME', 'STRING'];
property_tags[0x3A12] = ['ORIGINAL_ENTRYID', 'BINARY'];
property_tags[0x3A13] = ['ORIGINAL_DISPLAY_NAME', 'STRING'];
property_tags[0x3A14] = ['ORIGINAL_SEARCH_KEY', 'BINARY'];
property_tags[0x3A15] = ['POSTAL_ADDRESS', 'STRING'];
property_tags[0x3A16] = ['COMPANY_NAME', 'STRING'];
property_tags[0x3A17] = ['TITLE', 'STRING'];
property_tags[0x3A18] = ['DEPARTMENT_NAME', 'STRING'];
property_tags[0x3A19] = ['OFFICE_LOCATION', 'STRING'];
property_tags[0x3A1A] = ['PRIMARY_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A1B] = ['BUSINESS2_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A1C] = ['MOBILE_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A1D] = ['RADIO_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A1E] = ['CAR_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A1F] = ['OTHER_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A20] = ['TRANSMITABLE_DISPLAY_NAME', 'STRING'];
property_tags[0x3A21] = ['PAGER_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A22] = ['USER_CERTIFICATE', 'BINARY'];
property_tags[0x3A23] = ['PRIMARY_FAX_NUMBER', 'STRING'];
property_tags[0x3A24] = ['BUSINESS_FAX_NUMBER', 'STRING'];
property_tags[0x3A25] = ['HOME_FAX_NUMBER', 'STRING'];
property_tags[0x3A26] = ['COUNTRY', 'STRING'];
property_tags[0x3A27] = ['LOCALITY', 'STRING'];
property_tags[0x3A28] = ['STATE_OR_PROVINCE', 'STRING'];
property_tags[0x3A29] = ['STREET_ADDRESS', 'STRING'];
property_tags[0x3A2A] = ['POSTAL_CODE', 'STRING'];
property_tags[0x3A2B] = ['POST_OFFICE_BOX', 'STRING'];
property_tags[0x3A2C] = ['TELEX_NUMBER', 'STRING'];
property_tags[0x3A2D] = ['ISDN_NUMBER', 'STRING'];
property_tags[0x3A2E] = ['ASSISTANT_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A2F] = ['HOME2_TELEPHONE_NUMBER', 'STRING'];
property_tags[0x3A30] = ['ASSISTANT', 'STRING'];
property_tags[0x3A40] = ['SEND_RICH_INFO', 'BOOLEAN'];
property_tags[0x3A41] = ['WEDDING_ANNIVERSARY', 'SYSTIME'];
property_tags[0x3A42] = ['BIRTHDAY', 'SYSTIME'];
property_tags[0x3A43] = ['HOBBIES', 'STRING'];
property_tags[0x3A44] = ['MIDDLE_NAME', 'STRING'];
property_tags[0x3A45] = ['DISPLAY_NAME_PREFIX', 'STRING'];
property_tags[0x3A46] = ['PROFESSION', 'STRING'];
property_tags[0x3A47] = ['PREFERRED_BY_NAME', 'STRING'];
property_tags[0x3A48] = ['SPOUSE_NAME', 'STRING'];
property_tags[0x3A49] = ['COMPUTER_NETWORK_NAME', 'STRING'];
property_tags[0x3A4A] = ['CUSTOMER_ID', 'STRING'];
property_tags[0x3A4B] = ['TTYTDD_PHONE_NUMBER', 'STRING'];
property_tags[0x3A4C] = ['FTP_SITE', 'STRING'];
property_tags[0x3A4D] = ['GENDER', 'I2'];
property_tags[0x3A4E] = ['MANAGER_NAME', 'STRING'];
property_tags[0x3A4F] = ['NICKNAME', 'STRING'];
property_tags[0x3A50] = ['PERSONAL_HOME_PAGE', 'STRING'];
property_tags[0x3A51] = ['BUSINESS_HOME_PAGE', 'STRING'];
property_tags[0x3A52] = ['CONTACT_VERSION', 'CLSID'];
property_tags[0x3A53] = ['CONTACT_ENTRYIDS', 'MV_BINARY'];
property_tags[0x3A54] = ['CONTACT_ADDRTYPES', 'MV_STRING'];
property_tags[0x3A55] = ['CONTACT_DEFAULT_ADDRESS_INDEX', 'I4'];
property_tags[0x3A56] = ['CONTACT_EMAIL_ADDRESSES', 'MV_STRING'];
property_tags[0x3A57] = ['COMPANY_MAIN_PHONE_NUMBER', 'STRING'];
property_tags[0x3A58] = ['CHILDRENS_NAMES', 'MV_STRING'];
property_tags[0x3A59] = ['HOME_ADDRESS_CITY', 'STRING'];
property_tags[0x3A5A] = ['HOME_ADDRESS_COUNTRY', 'STRING'];
property_tags[0x3A5B] = ['HOME_ADDRESS_POSTAL_CODE', 'STRING'];
property_tags[0x3A5C] = ['HOME_ADDRESS_STATE_OR_PROVINCE', 'STRING'];
property_tags[0x3A5D] = ['HOME_ADDRESS_STREET', 'STRING'];
property_tags[0x3A5E] = ['HOME_ADDRESS_POST_OFFICE_BOX', 'STRING'];
property_tags[0x3A5F] = ['OTHER_ADDRESS_CITY', 'STRING'];
property_tags[0x3A60] = ['OTHER_ADDRESS_COUNTRY', 'STRING'];
property_tags[0x3A61] = ['OTHER_ADDRESS_POSTAL_CODE', 'STRING'];
property_tags[0x3A62] = ['OTHER_ADDRESS_STATE_OR_PROVINCE', 'STRING'];
property_tags[0x3A63] = ['OTHER_ADDRESS_STREET', 'STRING'];
property_tags[0x3A64] = ['OTHER_ADDRESS_POST_OFFICE_BOX', 'STRING'];
property_tags[0x3D00] = ['STORE_PROVIDERS', 'BINARY'];
property_tags[0x3D01] = ['AB_PROVIDERS', 'BINARY'];
property_tags[0x3D02] = ['TRANSPORT_PROVIDERS', 'BINARY'];
property_tags[0x3D04] = ['DEFAULT_PROFILE', 'BOOLEAN'];
property_tags[0x3D05] = ['AB_SEARCH_PATH', 'MV_BINARY'];
property_tags[0x3D06] = ['AB_DEFAULT_DIR', 'BINARY'];
property_tags[0x3D07] = ['AB_DEFAULT_PAB', 'BINARY'];
property_tags[0x3D09] = ['SERVICE_NAME', 'STRING'];
property_tags[0x3D0A] = ['SERVICE_DLL_NAME', 'STRING'];
property_tags[0x3D0B] = ['SERVICE_ENTRY_NAME', 'STRING'];
property_tags[0x3D0C] = ['SERVICE_UID', 'BINARY'];
property_tags[0x3D0D] = ['SERVICE_EXTRA_UIDS', 'BINARY'];
property_tags[0x3D0E] = ['SERVICES', 'BINARY'];
property_tags[0x3D0F] = ['SERVICE_SUPPORT_FILES', 'MV_STRING'];
property_tags[0x3D10] = ['SERVICE_DELETE_FILES', 'MV_STRING'];
property_tags[0x3D11] = ['AB_SEARCH_PATH_UPDATE', 'BINARY'];
property_tags[0x3D12] = ['PROFILE_NAME', 'STRING'];
property_tags[0x3E00] = ['IDENTITY_DISPLAY', 'STRING'];
property_tags[0x3E01] = ['IDENTITY_ENTRYID', 'BINARY'];
property_tags[0x3E02] = ['RESOURCE_METHODS', 'I4'];
// Service provider type
property_tags[0x3E03] = ['RESOURCE_TYPE', 'I4'];
property_tags[0x3E04] = ['STATUS_CODE', 'I4'];
property_tags[0x3E05] = ['IDENTITY_SEARCH_KEY', 'BINARY'];
property_tags[0x3E06] = ['OWN_STORE_ENTRYID', 'BINARY'];
property_tags[0x3E07] = ['RESOURCE_PATH', 'STRING'];
property_tags[0x3E08] = ['STATUS_STRING', 'STRING'];
property_tags[0x3E09] = ['X400_DEFERRED_DELIVERY_CANCEL', 'BOOLEAN'];
property_tags[0x3E0A] = ['HEADER_FOLDER_ENTRYID', 'BINARY'];
property_tags[0x3E0B] = ['REMOTE_PROGRESS', 'I4'];
property_tags[0x3E0C] = ['REMOTE_PROGRESS_TEXT', 'STRING'];
property_tags[0x3E0D] = ['REMOTE_VALIDATE_OK', 'BOOLEAN'];
property_tags[0x3F00] = ['CONTROL_FLAGS', 'I4'];
property_tags[0x3F01] = ['CONTROL_STRUCTURE', 'BINARY'];
property_tags[0x3F02] = ['CONTROL_TYPE', 'I4'];
property_tags[0x3F03] = ['DELTAX', 'I4'];
property_tags[0x3F04] = ['DELTAY', 'I4'];
property_tags[0x3F05] = ['XPOS', 'I4'];
property_tags[0x3F06] = ['YPOS', 'I4'];
property_tags[0x3F07] = ['CONTROL_ID', 'BINARY'];
property_tags[0x3F08] = ['INITIAL_DETAILS_PANE', 'I4'];

abstract class FixedLengthValueLoader {
    fixed_length: string = "";
    abstract load(value: any): any;
}

abstract class VariableLengthValueLoader {
    variable_length: string = "";
    abstract load(value: any): any;
}

class NULL extends FixedLengthValueLoader {
    load(value: any): any {
        // value is an eight-byte long bytestring with unused content.
        return null;
    }
}

class BOOLEAN extends FixedLengthValueLoader {
    load(value: any): any {
        // value is an eight-byte long bytestring holding a two-byte integer.
        return value[0] == 1;
    }
}

class INTEGER16 extends FixedLengthValueLoader {
    load(value: any): any {
        // value is an eight-byte long bytestring holding a two-byte integer.
        return value.slice(0, 2).reverse().reduce((a: any, b: any) => (a << 8) + b);
    }
}

class INTEGER32 extends FixedLengthValueLoader {
    load(value: any): any {
        // value is an eight-byte long bytestring holding a four-byte integer.
        return value.slice(0, 4).reverse().reduce((a: any, b: any) => (a << 8) + b);
    }
}

class INTEGER64 extends FixedLengthValueLoader {
    load(value: any): any {
        // value is an eight-byte long bytestring holding an eight-byte integer.
        return value.slice().reverse().reduce((a: any, b: any) => (bigInt(a).shiftLeft(8)).add(bigInt(b)));
    }
}

class INTTIME extends FixedLengthValueLoader {
    load(value: any): any {
        // value is an eight-byte long bytestring encoding the integer number of
        // 100-nanosecond intervals since January 1, 1601.
        //
        // Use bigint due to number type being too small to fit a 64-bit integer
        let delta: any = value.slice().reverse().reduce((a: any, b: any) => (bigInt(a).shiftLeft(8)).add(bigInt(b)));
        return new Date(new Date('1601-01-01T00:00:00Z').getTime() + Number(delta.divide(10000)));
    }
}

class BINARY extends VariableLengthValueLoader {
    load(value: any): any {
        return value;
    }
}

class STRING8 extends VariableLengthValueLoader {
    load(value: any): any {
        return new TextDecoder("utf-8").decode(new Uint8Array(value || []));
    }
}

class UNICODE extends VariableLengthValueLoader {
    load(value: any): any {
        return new TextDecoder("utf-16le").decode(new Uint8Array(value || []));
    }
}

class EMBEDDED_MESSAGE {
    async load(cfb: any, entry_name: string) {
        return await load_message_stream(cfb, entry_name, false);
    }
}


var property_types = new Array(0x102 + 1);
property_types[0x1] = new NULL();
property_types[0x2] = new INTEGER16();
property_types[0x3] = new INTEGER32();
property_types[0x4] = "FLOAT";
property_types[0x5] = "DOUBLE";
property_types[0x6] = "CURRENCY";
property_types[0x7] = "APPTIME";
property_types[0xa] = "ERROR";
property_types[0xb] = new BOOLEAN();
property_types[0xd] = new EMBEDDED_MESSAGE();
property_types[0x14] = new INTEGER64();
property_types[0x1e] = new STRING8();
property_types[0x1f] = new UNICODE();
property_types[0x40] = new INTTIME();
property_types[0x48] = "CLSID";
property_types[0xFB] = "SVREID";
property_types[0xFD] = "SRESTRICT";
property_types[0xFE] = "ACTIONS";
property_types[0x102] = new BINARY();

async function parse_properties(cfb: any, entry_name: string, is_top_level: boolean) {
    // Load stream content
    let entry = CFB.find(cfb, entry_name);
    if (entry == null) {
        return {};
    }

    // Skip header.
    let i = is_top_level ? 32 : 24;

    // Read 16-byte entries.
    let ret: any = {};
    while (i < entry.size) {
        // Read the entry
        let property_type = entry.content.slice(i+0, i+2);
        let property_tag = entry.content.slice(i+2, i+4);
        let flags = entry.content.slice(i+4, i+8);
        let value = entry.content.slice(i+8, i+16);
        i += 16;

        // Turn the byte strings into numbers and look up the property type
        property_type = property_type[0] + (property_type[1] << 8);
        property_tag = property_tag[0] + (property_tag[1] << 8);
        if (property_tag > property_tags.length || !property_tags[property_tag]) continue;
        let tag_name = property_tags[property_tag][0];
        let tag_type = property_types[property_type];

        if (tag_type instanceof FixedLengthValueLoader) {
            ret[tag_name] = (<FixedLengthValueLoader>tag_type).load(value);
        } else if (tag_type instanceof VariableLengthValueLoader) {
            // Look up the stream in the document that holds the value.
            let stream_name = "__substg1.0_" + 
                (<string>(property_tag.toString(16))).toUpperCase().padStart(4, '0') +
                (<string>(property_type.toString(16))).toUpperCase().padStart(4, '0');
            stream_name = entry_name.substring(0, entry_name.lastIndexOf('/')) + '/' + stream_name;
            value = CFB.find(cfb, stream_name);
            if (!value) continue;
            ret[tag_name] = (<VariableLengthValueLoader>tag_type).load(value.content);
        } else if (tag_type instanceof EMBEDDED_MESSAGE) {
           // Look up the stream in the document that holds the value.
           let stream_name = "__substg1.0_" + 
                (<string>(property_tag.toString(16))).toUpperCase().padStart(4, '0') +
                (<string>(property_type.toString(16))).toUpperCase().padStart(4, '0');
            stream_name = entry_name.substring(0, entry_name.lastIndexOf('/')) + '/' + stream_name;
            value = await (<EMBEDDED_MESSAGE>tag_type).load(cfb, stream_name);
            ret[tag_name] = value;
        }
    }

    return ret;
}

async function process_attachment(cfb: any, entry_name: string, msg: any) {
    // Load attachment stream
    let props = await parse_properties(cfb, entry_name + "/__properties_version1.0", false);

    // The attachment content
    let blob = props["ATTACH_DATA_BIN"];

    if (!blob) {
        return;
    }

    // Get the filename and mime type
    let filename = props["ATTACH_LONG_FILENAME"] || props["ATTACH_FILENAME"];

    // Determine the correct filename for embedded e-mails
    if (!filename) {
        if ("ATTACH_MIME_TAG" in props && props["ATTACH_MIME_TAG"] == "message/rfc822") {
            if ("DISPLAY_NAME" in props && props["DISPLAY_NAME"]) {
                filename = props["DISPLAY_NAME"].replace(/[/\\?%*:|"<>]/g, '-') + ".eml";
            } else {
                cfb.unknown_attachment_count = cfb.unknown_attachment_count || 0;
                cfb.unknown_attachment_count++;
                filename = "unknown_" + cfb.unknown_attachment_count + ".eml"
            }
            props["ATTACH_MIME_TAG"] = "message/rfc822";
        } else {
            cfb.unknown_attachment_count = cfb.unknown_attachment_count || 0;
            cfb.unknown_attachment_count++;
            filename = "unknown_" + cfb.unknown_attachment_count + ".dat"
        }
    }

    let mime_type = props["ATTACH_MIME_TAG"] || "application/octet-stream";
    filename = filename.split("/").slice(-1)[0].split("\\").slice(-1)[0]

    msg.attachments.push({
        name: filename,
        contentType: mime_type,
        data: Buffer.from(blob)
    });
}

async function load_message_stream(cfb: any, entry_name: string, is_top_level: boolean): Promise<any> {
    // Load stream data
    let props = await parse_properties(cfb, entry_name + "/__properties_version1.0", is_top_level);

    // Construct the MIME message
    let msg: any = {}

    // Add the raw headers, if known.
    let headers_obj: any = {};

    if ('TRANSPORT_MESSAGE_HEADERS' in props) {
        let headers = props['TRANSPORT_MESSAGE_HEADERS'];
        (<string>headers).split("\r\n")
                        .filter(h => h.indexOf(': ') >= 0)
                        .map(h => [h.substring(0, h.indexOf(': ')), h.substring(h.indexOf(': ') + 2)])
                        .filter(h => h[0] != "Content-Type")
                        .forEach(h => headers_obj[h[0]] = h[1]);
    } else {
        // Construct common headers from metadata.
        if ("MESSAGE_DELIVERY_TIME" in props) {
            headers_obj["Date"] = moment(props["MESSAGE_DELIVERY_TIME"]).format("ddd, DD MMM YYYY HH:mm:ss ZZ")
        }
        if ("SENDER_NAME" in props && props["SENDER_NAME"]) {
            if ("SENT_REPRESENTING_NAME" in props && props["SENT_REPRESENTING_NAME"] &&
                props["SENDER_NAME"] != props["SENT_REPRESENTING_NAME"]) {
                props["SENDER_NAME"] = props["SENDER_NAME"] + " (" + props["SENT_REPRESENTING_NAME"] + ")";
            }
            headers_obj["From"] = props["SENDER_NAME"];
        }
        if ("SENDER_EMAIL_ADDRESS" in props && props["SENDER_EMAIL_ADDRESS"]) {
            headers_obj["From"] = (headers_obj["From"] || "") + " <" + props["SENDER_EMAIL_ADDRESS"] + ">";
        }
        if ("DISPLAY_TO" in props && props["DISPLAY_TO"]) {
            headers_obj["To"] = (<string>props["DISPLAY_TO"]).replace(/\x00$/, "");
        }
        if ("DISPLAY_CC" in props && props["DISPLAY_CC"]) {
            headers_obj["CC"] = props["DISPLAY_CC"];
        }
        if ("DISPLAY_BCC" in props && props["DISPLAY_BCC"]) {
            headers_obj["BCC"] = props["DISPLAY_BCC"];
        }
        if ("SUBJECT" in props && props["SUBJECT"]) {
            headers_obj["Subject"] = props["SUBJECT"];
        }
    }

    // Add the plain-text body from the BODY field.
    let attachment_refs: any = {};
    if ("BODY" in props && !("RTF_COMPRESSED" in props)) {
        msg["text"] = props["BODY"]
    } else {
        // Decompress the RTF and then deencapsulate the RTF to obtain the original
        // HTML representation of the email
        let rtf = new Uint8Array(decompressRTF(props["RTF_COMPRESSED"]));

        // Check if the RTF actually contains HTML tags, otherwise use plaintext
        if (new TextDecoder("utf-8").decode(rtf).indexOf("\\*\\htmltag") >= 0) {
            let html: string = <string>deEncapsulateSync(Buffer.from(rtf), { decode: iconvLite.decode }).text;
            msg["html"] = html;

            // Detect all the inlined-attachments
            let r = /src="cid:(([^@]+)@[A-Z0-9]+\.[A-Z0-9]+)"/g;
            let m;
            while ((m = r.exec(html)) !== null) {
                attachment_refs[m[2]] = m[1];
            }
        } else {
            msg["text"] = props["BODY"];
        }
    }

    msg["headers"] = headers_obj;

    // Copy all attachments
    msg.attachments = msg.attachments || [];
    for (let i = 0; i < cfb.FullPaths.length; i++) {
        if (cfb.FullPaths[i].indexOf("/__attach_version1.0_#") >= 0 &&
            cfb.FullPaths[i].indexOf(entry_name.replace(/\/$/g, "")) == 0 &&
            cfb.FullPaths[i].replace(/\/$/g, "") != entry_name.replace(/\/$/g, "") &&
            cfb.FullPaths[i].replace(/\/$/g, "").split('/').length == entry_name.replace(/\/$/g, "").split('/').length + 1) {
            await process_attachment(cfb, (<string>cfb.FullPaths[i]).replace(/\/$/g, ""), msg);
        }
    }

    // Fix inline-attachment references
    msg.attachments.forEach((a: any) => {
        if (a.name in attachment_refs) {
            a.cid = attachment_refs[a.name];
        }
    });

    return new Promise((resolve, reject) => {
        eml_format.build(msg, function(error: any, eml: any) {
            if (error) {
                reject(error);
            } else {
                resolve(eml);
            }
        });
    });
}

moment.locale("en");

export async function msg2eml(obj: any) {
    let arrBuffer: Uint8Array;
    if (obj.constructor && obj.constructor.name === 'Blob') {
        arrBuffer = new Uint8Array(await obj.arrayBuffer());
    } else if ((obj.constructor && obj.constructor.name == 'Array') ||
               (obj.constructor && obj.constructor.name == 'ArrayBuffer')) {
        arrBuffer = new Uint8Array(obj);
    } else {
        throw 'Unknown or unsupported source type: ' + (obj.constructor ? obj.constructor.name : obj);
    }
    
    let result = CFB.parse(arrBuffer);
    return await load_message_stream(result, "Root Entry", true);
};

if (typeof window !== 'undefined' && window) {
    (<any>window).msg2eml = msg2eml;
}

if (typeof module !== 'undefined' && module && module.exports) {
    module.exports = msg2eml;
}
