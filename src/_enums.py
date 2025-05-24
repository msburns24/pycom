from enum import Enum


class OlAddressEntryUserType(Enum):
    '''
    Represents the type of user for the `AddressEntry` or object derived from
    `AddressEntry`.

    - `EXCHANGE_USER: 0` - An Exchange user that belongs to the same Exchange
      forest.
    - `EXCHANGE_DISTRIBUTION_LIST: 1` - An address entry that is an Exchange
      distribution list.
    - `EXCHANGE_PUBLIC_FOLDER: 2` - An address entry that is an Exchange public
      folder.
    - `EXCHANGE_AGENT: 3` - An address entry that is an Exchange agent.
    - `EXCHANGE_ORGANIZATION: 4` - An address entry that is an Exchange
      organization.
    - `EXCHANGE_REMOTE_USER: 5` - An Exchange user that belongs to a different
      Exchange forest.
    - `OUTLOOK_CONTACT: 10` -  An address entry in an Outlook Contacts folder.
    - `OUTLOOK_DISTRIBUTION_LIST: 11` -  An address entry that is an Outlook
      distribution list.
    - `LDAP: 20` -  An address entry that uses the Lightweight Directory Access
      Protocol (LDAP).
    - `SMTP: 30` -  An address entry that uses the Simple Mail Transfer
      Protocol (SMTP).
    - `OTHER: 40` -  A custom or some other type of address entry such as FAX.
    '''
    EXCHANGE_USER               = 0
    EXCHANGE_DISTRIBUTION_LIST  = 1
    EXCHANGE_PUBLIC_FOLDER      = 2
    EXCHANGE_AGENT              = 3
    EXCHANGE_ORGANIZATION       = 4
    EXCHANGE_REMOTE_USER        = 5
    OUTLOOK_CONTACT             = 10
    OUTLOOK_DISTRIBUTION_LIST   = 11
    LDAP                        = 20
    SMTP                        = 30
    OTHER                       = 40
    pass


class OlAutoDiscoverConnectionMode(Enum):
    '''
    Specifies the type of connection to the Exchange server for the
    auto-discovery service.

    - `UNKNOWN: 0` - Other/unknown connection or no connection.
    - `EXTERNAL: 1` -  Connection is over the Internet.
    - `INTERNAL: 2` -  Connection is over the Intranet.
    - `INTERNAL_DOMAIN: 3` -  Connection is same domain over the Intranet.
    '''
    UNKNOWN         = 0
    EXTERNAL        = 1
    INTERNAL        = 2
    INTERNAL_DOMAIN = 3


class OlAccountType(Enum):
    '''
    Specifies the type of an Account.
    
    - `EXCHANGE: 0` - An Exchange account.
    - `IMAP: 1` - An IMAP account.
    - `POP3: 2` - A POP3 account.
    - `HTTP: 3` - An HTTP account.
    - `EAS: 4` - An account using Exchange ActiveSync (EAS).
    - `OTHER_ACCOUNT: 5` - Other/unknown account.
    '''
    EXCHANGE       = 0
    IMAP           = 1
    POP3           = 2
    HTTP           = 3
    EAS            = 4
    OTHER_ACCOUNT  = 5


class OlDefaultFolders(Enum):
    '''
    Specifies the folder type for the current Microsoft Outlook profile.

    - `DELETED_ITEMS: 3` - The Deleted Items folder.
    - `OUTBOX: 4` - The Outbox folder.
    - `SENT_MAIL: 5` - The Sent Mail folder.
    - `INBOX: 6` - The Inbox folder.
    - `CALENDAR: 9` - The Calendar folder.
    - `CONTACTS: 10` - The Contacts folder.
    - `JOURNAL: 11` - The Journal folder.
    - `NOTES: 12` - The Notes folder.
    - `TASKS: 13` - The Tasks folder.
    - `DRAFTS: 16` - The Drafts folder.
    - `ALL_PUBLIC_FOLDERS: 18` - The All Public Folders folder in the Exchange
       Public Folders store. Only available for an Exchange account.
    - `CONFLICTS: 19` - The Conflicts folder (subfolder of Sync Issues folder).
       Only available for an Exchange account.
    - `SYNC_ISSUES: 20` - The Sync Issues folder. Only available for an
       Exchange account.
    - `LOCAL_FAILURES: 21` - The Local Failures folder (subfolder of Sync
       Issues folder). Only available for an Exchange account.
    - `SERVER_FAILURES: 22` - The Server Failures folder (subfolder of Sync
       Issues folder). Only available for an Exchange account.
    - `JUNK: 23` - The Junk E-Mail folder.
    - `RSS_FEEDS: 25` - The RSS Feeds folder.
    - `TO_DO: 28` - The To Do folder.
    - `MANAGED_EMAIL: 29` - The top-level folder in the Managed Folders group.
       For more information on Managed Folders, see Help in Outlook. Only
       available for an Exchange account.
    - `SUGGESTED_CONTACTS: 30` - The Suggested Contacts folder.
    '''
    DELETED_ITEMS       = 3
    OUTBOX              = 4
    SENT_MAIL           = 5
    INBOX               = 6
    CALENDAR            = 9
    CONTACTS            = 10
    JOURNAL             = 11
    NOTES               = 12
    TASKS               = 13
    DRAFTS              = 16
    ALL_PUBLIC_FOLDERS  = 18
    CONFLICTS           = 19
    SYNC_ISSUES         = 20
    LOCAL_FAILURES      = 21
    SERVER_FAILURES     = 22
    JUNK                = 23
    RSS_FEEDS           = 25
    TO_DO               = 28
    MANAGED_EMAIL       = 29
    SUGGESTED_CONTACTS  = 30



class OlExchangeConnectionMode(Enum):
    '''
    Specifies whether the account is connected to an Exchange server and if so,
    the connection mode.
    
    - `NO_EXCHANGE: 0` - The account does not use an Exchange server.
    - `OFFLINE: 100` - The account is not connected to an Exchange server and
       is in the classic offline mode. This also occurs when the user selects
       Work Offline from the File menu.
    - `CACHED_OFFLINE: 200` - The account is using cached Exchange mode and the
      user has selected Work Offline from the File menu.
    - `DISCONNECTED: 300` - The account has a disconnected connection to the
      Exchange server.
    - `CACHED_DISCONNECTED: 400` - The account is using cached Exchange mode
      with a disconnected connection to the Exchange server.
    - `CACHED_CONNECTED_HEADERS: 500` - The account is using cached Exchange
      mode on a dial-up or slow connection with the Exchange server, such that
      only headers are downloaded. Full item bodies and attachments remain on
      the server. The user can also select this state manually regardless of
      connection speed.
    - `CACHED_CONNECTED_DRIZZLE: 600` - The account is using cached Exchange
      mode such that headers are downloaded first, followed by the bodies and
      attachments of full items.
    - `CACHED_CONNECTED_FULL: 700` - The account is using cached Exchange mode
      on a Local Area Network or a fast connection with the Exchange server.
      The user can also select this state manually, disabling auto-detect logic
      and always downloading full items regardless of connection speed.
    - `ONLINE: 800` - The account is connected to an Exchange server and is in
      the classic online mode.
    '''
    NO_EXCHANGE               = 0
    OFFLINE                   = 100
    CACHED_OFFLINE            = 200
    DISCONNECTED              = 300
    CACHED_DISCONNECTED       = 400
    CACHED_CONNECTED_HEADERS  = 500
    CACHED_CONNECTED_DRIZZLE  = 600
    CACHED_CONNECTED_FULL     = 700
    ONLINE                    = 800


class OlFolderDisplayMode(Enum):
    '''
    Specifies the folder display mode.

    - `NORMAL: 0` - Folder is displayed with navigation pane on the left and
      folder contents on the right.
    - `FOLDER_ONLY: 1` - Only the contents of the selected folder are
      displayed.
    - `NO_NAVIGATION: 2` - Folder contents are displayed but no navigation pane
      is shown.
    '''
    NORMAL         = 0
    FOLDER_ONLY    = 1
    NO_NAVIGATION  = 2


class OlItemType(Enum):
    '''
    Indicates the Outlook Item type.

    - `MAIL_ITEM: 0` - Represents a `MailItem`
    - `APPOINTMENT_ITEM: 1` - Represents an `AppointmentItem`
    - `CONTACT_ITEM: 2` - Represents a `ContactItem`
    - `TASK_ITEM: 3` - Represents a `TaskItem`
    - `JOURNAL_ITEM: 4` - Represents a `JournalItem`
    - `NOTE_ITEM: 5` - Represents a `NoteItem`
    - `POST_ITEM: 6` - Represents a `PostItem`
    - `DISTRIBUTION_LIST_ITEM: 7` - Represents an `DistListItem`
    - `MOBILE_ITEM_SMS: 11` - A `MobileItem` object that is a Short Message
      Service (SMS) message.
    - `MOBILE_ITEM_MMS: 12` - A `MobileItem` object that is a Multimedia
      Messaging Service (MMS) message.
    '''
    MAIL_ITEM               = 0
    APPOINTMENT_ITEM        = 1
    CONTACT_ITEM            = 2
    TASK_ITEM               = 3
    JOURNAL_ITEM            = 4
    NOTE_ITEM               = 5
    POST_ITEM               = 6
    DISTRIBUTION_LIST_ITEM  = 7
    MOBILE_ITEM_SMS         = 11
    MOBILE_ITEM_MMS         = 12


class OlSharingProvider(Enum):
    '''
    Indicates the sharing provider associated with a `SharingItem` object.

    - `UNKNOWN 0` - Represents an unknown sharing provider. This value is used
      if the sharing provider GUID in the sharing message does not match the
      GUID of any of the sharing providers represented in this enumeration.
    - `EXCHANGE 1` - Represents the Exchange sharing provider.
    - `WEB_CAL 2` - Represents the WebCal sharing provider.
    - `PUB_CAL 3` - Represents the PubCal sharing provider.
    - `ICAL 4` - Represents the iCalendar sharing provider.
    - `SHAREPOINT 5` - Represents the Microsoft SharePoint Foundation sharing
      provider.
    - `RSS 6` - Represents the Really Simple Syndication (RSS) sharing
      provider.
    - `FEDERATE 7` - Represents a federated sharing provider. A SharingItem
      object with this type of provider is used for sharing relationships
      across organizational boundaries (for example, between two organizations
      using Microsoft Exchange Server 2010).
    '''
    UNKNOWN     = 0
    EXCHANGE    = 1
    WEB_CAL     = 2
    PUB_CAL     = 3
    ICAL        = 4
    SHAREPOINT  = 5
    RSS         = 6
    FEDERATE    = 7


class OlShowItemCount(Enum):
    '''
    Indicates which type of count for Microsoft Outlook items is displayed for
    folders in the Outlook Navigation Pane.

    - `NO_ITEM_COUNT: 0` - No item count displayed.
    - `SHOW_UNREAD_ITEM_COUNT: 1` - Shows count of unread items.
    - `SHOW_TOTAL_ITEM_COUNT: 2` - Shows count of total number of items.
    '''
    NO_ITEM_COUNT           = 0
    SHOW_UNREAD_ITEM_COUNT  = 1
    SHOW_TOTAL_ITEM_COUNT   = 2


class OlStorageIdentifierType(Enum):
    '''
    Specifies the type of identifier for a `StorageItem` object.

    - `SUBJECT: 0` - Identifies a `StorageItem` by `subject`.
    - `ENTRY_ID: 1` - Identifies a `StorageItem` by `entry_id`.
    - `MESSAGE_CLASS: 2` - Identifies a `StorageItem` by message class.
    '''
    SUBJECT        = 0
    ENTRY_ID       = 1
    MESSAGE_CLASS  = 2


class OlTableContents(Enum):
    '''
    Specifies the type of items in a folder.
    
    - `USER_ITEMS: 0` - Only the non-hidden user items in the folder.
    - `HIDDEN_ITEMS: 1` - Only the hidden items in the folder.
    '''
    USER_ITEMS    = 0
    HIDDEN_ITEMS  = 1
