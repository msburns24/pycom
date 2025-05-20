from __future__ import annotations
from typing import Optional
from rich import print, inspect
from win32com import client
from win32com.client import Dispatch, CDispatch


# Constants
INBOX_FOLDER_NUMBER = 6


def main() -> int:
    app = Application()
    namespace = app.get_namespace('MAPI')
    inbox = namespace.get_default_folder(INBOX_FOLDER_NUMBER)
    inspect(inbox)
    return 0


def extract_attributes(
        to_object: object,
        from_object: object,
        attrs_map: dict[str, str],
) -> None:
    '''
    Extracts attributes specified in `attrs_map` and adds them to the provided
    object. When an exception is raised, stores `None`.

    Parameters
    ----------
    to_object : object
        The object in which to store the extracted attributes.
    from_object : object
        The object from which to extract the attributes
    attrs_map : dict[str, str]
        A map of `from_attr_name` to `to_attr_name`
    
    Returns
    -------
    None
    '''
    for from_attr_name, to_attr_name in attrs_map.items():
        try:
            value = getattr(from_object, from_attr_name)
        except:
            value = None
        setattr(to_object, to_attr_name, value)
    return


class Application:
    '''
    Represents the entire Outlook application.

    Attributes
    ----------
    active_explorer : CDispatch
        Returns the topmost `Explorer` object on the desktop.
    active_window : CDispatch
        Returns an object representing the topmost Microsoft Outlook window on
        the desktop, either an `Explorer` or an `Inspector` object.
    assistance : CDispatch
        Returns an `IAssistance`
    com_add_ins : CDispatch
        Returns a `COMAddIns` collection that represents all the Component
        Object Model (COM) add-ins currently loaded in Microsoft Outlook.
    data_privacy_options : CDispatch
        Data privacy options (no documentation available)
    default_profile_name : str
        Returns a string representing the name of the default profile name.
        Read-only.
    explorers : CDispatch
        Returns an `Explorers` collection object that contains the `Explorer`
        objects representing all open explorers. Read-only.
    inspectors : CDispatch
        Returns an `Inspectors` collection object that contains the `Inspector`
        objects representing all open inspectors. Read-only.
    is_trusted : bool
        Returns a boolean to indicate if an add-in or external caller is
        considered trusted by Outlook. Read-only.
    language_settings : CDispatch
        Returns a `LanguageSettings`
    name : str
        Returns a string value that represents the display name for the object.
        Read-only.
    picker_dialog : CDispatch
        Returns a `PickerDialog` object that provides the functionality to
        select people or data in a dialog box. Read-only.
    product_code : str
        Returns a string specifying the Microsoft Outlook globally unique
        identifier (GUID)
    reminders : CDispatch
        Returns a `Reminders` collection that represents all current
        reminders. Read-only.
    session : CDispatch
        Returns the `NameSpace` object for the current session. Read-only.
    time_zones : CDispatch
        Returns a `TimeZones` collection that represents the set of time zones
        supported by Outlook. Read-only.
    version : str
        Returns or sets a string indicating the number of the version.
        Read-only. 
    
    Methods
    -------
    advanced_search(scope, filter, search_sub_folders, tag)
        Performs a search based on a specified DAV Searching and Locating
        (DASL) search string.
    copy_file(file_path, dest_folder_path)
        Copies a file from a specified location into a Micro1soft Outlook
        store.
    create_item(item_type)
        Creates and returns a new Microsoft Outlook item.
    create_item_from_template(template_path, in_folder=None)
        Creates a new Microsoft Outlook item from an Outlook template (`.oft`)
        and returns the new item.
    create_object(object_name)
        Creates an Automation object of the specified class.
    get_namespace(type_='MAPI')
        Returns a NameSpace object of the specified type.
    get_object_reference(item, reference_type)
        Creates a strong or weak object reference for a specified `Outlook`
        object.
    is_search_synchronous(look_in_folders)
        Returns a boolean indicating if a search will be synchronous or
        asynchronous.
    refresh_form_region_definition(region_name='')
        Refreshes the cache by obtaining the current definition from the
        Windows registry for one or all of the form regions that are defined
        for the local machine and the current user.
    '''

    def __init__(self) -> None:
        self._application = Dispatch('Outlook.Application')

        # Type hints
        self.assistance:            CDispatch|None  = None
        self.com_add_ins:           CDispatch|None  = None
        self.data_privacy_options:  CDispatch|None  = None
        self.default_profile_name:  str|None        = None
        self.explorers:             CDispatch|None  = None
        self.inspectors:            CDispatch|None  = None
        self.is_trusted:            bool|None       = None
        self.language_settings:     CDispatch|None  = None
        self.name:                  str|None        = None
        self.picker_dialog:         CDispatch|None  = None
        self.product_code:          str|None        = None
        self.reminders:             CDispatch|None  = None
        self.session:               CDispatch|None  = None
        self.time_zones:            CDispatch|None  = None
        self.version:               str|None        = None

        attrs_map = {
            'Assistance':          'assistance',
            'COMAddIns':           'com_add_ins',
            'DataPrivacyOptions':  'data_privacy_options',
            'DefaultProfileName':  'default_profile_name',
            'Explorers':           'explorers',
            'Inspectors':          'inspectors',
            'IsTrusted':           'is_trusted',
            'LanguageSettings':    'language_settings',
            'Name':                'name',
            'PickerDialog':        'picker_dialog',
            'ProductCode':         'product_code',
            'Reminders':           'reminders',
            'Session':             'session',
            'TimeZones':           'time_zones',
            'Version':             'version',
        }
        # for com_name, new_name in attrs_map.items():
        #     try:
        #         value = getattr(self._application, com_name)
        #         setattr(self, new_name, value)
        #     except:
        #         pass
        extract_attributes(self, self._application, attrs_map)
        return
    
    @property
    def active_explorer(self) -> CDispatch:
        '''Returns the topmost `Explorer` object on the desktop.'''
        return self._application.ActiveExplorer()
    
    @property
    def active_window(self) -> CDispatch:
        '''Returns an object representing the topmost Microsoft Outlook window
        on the desktop, either an `Explorer` or an `Inspector` object.'''
        return self._application.ActiveWindow()
    
    def advanced_search(
            self,
            scope: str,
            filter,
            search_sub_folders,
            tag,
    ) -> None:
        '''
        Performs a search based on a specified DAV Searching and Locating
        (DASL) search string.

        Parameters
        ----------
        scope: str
            The scope of the search. For example, the folder path of a folder.
            To specify multiple folder paths, enclose each folder path in
            single quotes and separate the single quoted folder paths with a
            comma.
        filter
            The DASL search filter that defines the parameters of the search.
        search_sub_folders
            Determines if the search will include any of the folder's
            subfolders.
        tag
            The name given as an identifier for the search.
        
        Returns
        -------
        search : Search
            A `Search` object that represents the results of the search.
        
        Remarks
        -------
        You can run multiple searches simultaneously by calling the
        `advanced_search` method in successive lines of code. However, you
        should be aware that programmatically creating a large number of search
        folders can result in significant simultaneous search activity that
        would affect the performance of Outlook, especially if Outlook conducts
        the search in online Exchange mode.

        The `advanced_search` method and related features in the Outlook object
        model do not create a Search Folder that will appear in the Outlook
        user interface. However, you can use the `save(string)` method of the
        `Search` object that is returned to create a Search Folder that will
        appear in the Search Folders list in the Outlook user interface.

        Using the `scope` parameter, you can specify one or more folders in the
        same store, but you may not specify multiple folders in multiple
        stores. To specify multiple folders for the `scope` parameter, use a
        comma character between each folder path and enclose each folder path
        in single quotes. For default folders such as Inbox or Sent Items, you
        can use the simple folder name instead of the full folder path.
        '''
        return self._application.AdvancedSearch(scope,
                                                filter,
                                                search_sub_folders,
                                                tag)
    
    def copy_file(self, file_path: str, dest_folder_path: str) -> CDispatch:
        '''
        Copies a file from a specified location into a Microsoft Outlook store.

        Parameters
        ----------
        file_path : str
            The path name of the object you want to copy.
        dest_folder_path : str
            The location you want to copy the file to.
        
        Returns
        -------
        obj : CDispatch
            An `Object` value that represents the copied file.
        '''
        return self._application.CopyFile(file_path, dest_folder_path)
    
    def create_item(self, item_type: CDispatch) -> CDispatch:
        '''
        Creates and returns a new Microsoft Outlook item.

        Parameters
        ----------
        item_type : CDispatch
            The Outlook item type for the new item.
        
        Returns
        -------
        object_ : CDispatch
            An Object value that represents the new Outlook item.

        Remarks
        -------
        The `create_item` method can only create default Outlook items. To
        create new items using a custom form, use the `add()` method on the
        `items` collection.
        '''
        return self._application.CreateItem(item_type)

    def create_item_from_template(
            self,
            template_path: str,
            in_folder: Optional[CDispatch]=None
    ) -> CDispatch:
        '''
        Creates a new Microsoft Outlook item from an Outlook template (`.oft`)
        and returns the new item.

        Parameters
        ----------
        template_path : str
            The path and file name of the Outlook template for the new item.
        in_folder : CDispatch
            The folder in which the item is to be created. If this argument is
            omitted, the default folder for the item type will be used.
        
        Returns
        -------
        object_ : CDispatch
            An Object value that represents the new Microsoft Outlook item.
        
        Remarks
        -------
        New items will always open in compose mode, as opposed to read mode,
        regardless of the mode in which the items were saved to disk.
        '''
        return self._application.CreateItemFromTemplate(template_path,
                                                        in_folder)
    
    def create_object(self, object_name: str) -> CDispatch:
        '''
        Creates an `Automation` object of the specified class.

        Parameters
        ----------
        object_name : str
            The class name of the object to create. For information about valid
            class names, see
            [OLE Programmatic Identifiers](http://go.microsoft.com/fwlink/?LinkId=87946).
        
        Returns
        -------
        object_ : CDispatch
            An Object value that represents the new `Automation` object
            instance. If the application is already running, `create_object`
            will create a new instance.

        Remarks
        -------
        This method is provided so that other applications can be automated
        from Microsoft Visual Basic Scripting Edition (VBScript) 1.0, which did
        not include a `create_object` method. `create_object` has been included
        in VBScript version 2.0 and later. This method should not be used to
        automate Microsoft Outlook from VBScript.
        '''
        return self._application.CreateObject(object_name)
    
    def get_namespace(self, type_: str='MAPI') -> NameSpace:
        '''
        Returns a NameSpace object of the specified type.

        Parameters
        ----------
        type_ : str
            The type of name space to return.
        
        Returns
        -------
        namespace : CDispatch
            A NameSpace object that represents the specified namespace.
        
        Remarks
        -------
        The only supported name space type is "MAPI". The GetNameSpace method
        is functionally equivalent to the Session property, which was
        introduced in Microsoft Outlook 98.
        '''
        return NameSpace(self, type_)
    
    def _get_namespace(self, type_: str='MAPI') -> CDispatch:
        '''
        Private method to access `pywin32` method.
        '''
        return self._application.GetNamespace(type_)
    
    def get_object_reference(
            self,
            item: CDispatch,
            reference_type: CDispatch
    ) -> CDispatch:
        '''
        Creates a strong or weak object reference for a specified `Outlook`
        object.

        Parameters
        ----------
        item : CDispatch
            The object from which to obtain a strong or weak object reference.
        reference_type : CDispatch
            The type of object reference.

        Returns
        -------
        object_ : CDispatch
            An `Object` that represents a strong or weak object reference for
            the specified object.
        
        Remarks
        -------
        This method returns a weak or strong object reference for the object
        specified in `item`.

        **Note:** Outlook can fail to close successfully if an add-in retains
        strong object references. Always dereference a strong object reference
        once it is no longer needed by the add-in.
        '''
        return self._application.GetObjectReference(item, reference_type)
    
    def is_search_synchronous(self, look_in_folders: str) -> bool:
        '''
        Returns a boolean indicating if a search will be synchronous or
        asynchronous.

        Parameters
        ----------
        look_in_folders : str
            The path name of the folders that the search will search through.

        Returns
        -------
        result : bool
            `True` if the search is synchronous; otherwise, `False`.

        Remarks
        -------
        If the search is synchronous, the `advanced_search` method will not
        return until the search has completed. Conversely, if the search is
        asynchronous, the `advanced_search` method will immediately return.
        In order to get meaningful results from an asynchronous search, use the
        `AdvancedSearchComplete` event to notify you when the search has
        finished.
        '''
        return self._application.IsSearchSynchronous(look_in_folders)
    
    def refresh_form_region_definition(self, region_name: str='') -> None:
        '''
        Refreshes the cache by obtaining the current definition from the
        Windows registry for one or all of the form regions that are defined
        for the local machine and the current user.

        Parameters
        ----------
        region_name : str
            The internal name of the form region whose definition you want to
            refresh in the cache. To refresh all form region definitions,
            specify an empty string.
        
        Returns
        -------
        None

        Remarks
        -------
        When Microsoft Outlook starts, it reads the Windows registry to obtain
        a list of form regions and their definitions, and then caches the data.
        The definitions are stored in the registry under the local machine key
        (as `HKEY_LOCAL_MACHINE/Software/Microsoft/Office/Outlook/FormRegions`)
        and under the current user key (as
        `HKEY_CURRENT_USER/Software/Microsoft/Office/Outlook/FormRegions`). The
        definitions describe the layout, behavior, and other characteristics of
        each form region. If you register a form region or modify the
        definition of a form region after Outlook starts, you can use the
        `refresh_form_region_definition` method to instruct Outlook to obtain
        the updated information.

        The `region_name` argument should match the `internal_name` property of
        the form region whose definition you are refreshing. The internal name
        of a form region supports only ASCII characters. If you specify an
        empty string, Outlook reads the Windows registry to obtain definitions
        for all of the form regions that are defined for the local machine and
        the current user.
        '''
        return self._application.RefreshFormRegionDefinition(region_name)
    

class NameSpace:
    '''
    Represents an abstract root object for any data source.

    Properties
    ----------
    accounts : CDispatch | None
        An `Accounts` collection object that represents all the `Account`
        objects in the current profile. Read-only.
    address_lists : CDispatch | None
        An `AddressLists` collection representing a collection of the address
        lists available for this session. Read-only.
    application : Application
        The parent Outlook application for the object.
    auto_discover_connection_mode : int | None
        A constant that specifies the type of connection to the Exchange
        server for auto-discovery service. Read-only.
    auto_discover_xml : str | None
        Information in XML retrieved from the auto-discovery service of an
        Exchange server. Read-only.
    categories : CDispatch | None
        The set of `Category` objects available to the namespace. Read/write.
    current_profile_name : str | None
        The name of the current profile. Read-only.
    current_user : CDispatch | None
        The display name of the currently logged-on user as a `Recipient`
        object. Read-only.
    default_store : CDispatch | None
        The default `Store` for the profile. Read-only.
    exchange_connection_mode : int | None
        A constant that indicates the current connection mode the user is
        using. Read-only.
    exchange_mailbox_server_name : str | None
        The name of the Exchange server on which the active mailbox is hosted.
        Read-only.
    exchange_mailbox_server_version : str | None
        The full version of the Exchange server on which the active mailbox is
        hosted. Read-only.
    folders : CDispatch | None
        All the folders contained in the specified NameSpace. Read-only.
    offline : bool | None
        Indicates `True` if Outlook is offline (not connected to an Exchange
        server), and `False` if online (connected to an Exchange server).
        Read-only.
    session : CDispatch | None
        The `NameSpace` object for the current session. Read-only.
    stores : CDispatch | None
        All the `Store` objects in the current profile. Read-only.
    sync_objects : CDispatch | None
        Contains all Send/Receive groups. Read-only.

    Methods
    -------
    add_store(store)
        Adds a Personal Folders (`.pst`) file to the current profile.
    add_store_ex(store, type_)
        Adds a Personal Folders file (`.pst`) in the specified format to the
        current profile.
    compare_entry_ids(first_entry_id, second_entry_id)
        Returns a boolean that indicates if two entry ID values refer to the
        same Outlook item.
    create_contact_card(address_entry)
        Creates an instance of a `ContactCard` object for the contact that is
        specified by the `address_entry` parameter.
    create_recipient(recipient_name)
        Creates a `Recipient` object.
    create_sharing_item(context, provider=None)
        Creates a new `SharingItem` object.
    dial(contact_item)
        Displays the New Call dialog box that allows users to dial the primary
        phone number of a specified contact.
    get_address_entry_from_id(id_)
        Returns an `AddressEntry` object that represents the address entry
        specified by ID.
    get_default_folder(folder_type)
        Returns a `Folder` object that represents the default folder of the
        requested type for the current profile; for example, obtains the
        default `Calendar` folder for the user who is currently logged on.
    GetFolderFromID
    GetGlobalAddressList
    GetIDsOfNames
    GetItemFromID
    GetRecipientFromID
    GetSelectNamesDialog
    GetSharedDefaultFolder
    GetStoreFromID
    GetTypeInfo
    GetTypeInfoCount
    Invoke
    Logoff
    Logon
    OpenSharedFolder
    OpenSharedItem
    PickFolder
    QueryInterface
    RefreshRemoteHeaders
    Release
    RemoveStore
    SendAndReceive
    '''

    def __init__(self, application: Application, type_: str='MAPI') -> None:
        self.application = application
        self._namespace = self.application._get_namespace(type_)

        # Type hints
        self.accounts:                         CDispatch|None  = None
        self.address_lists:                    CDispatch|None  = None
        self.auto_discover_connection_mode:    int|None        = None
        self.auto_discover_xml:                str|None        = None
        self.categories:                       CDispatch|None  = None
        self.current_profile_name:             str|None        = None
        self.current_user:                     CDispatch|None  = None
        self.default_store:                    CDispatch|None  = None
        self.exchange_connection_mode:         int|None        = None
        self.exchange_mailbox_server_name:     str|None        = None
        self.exchange_mailbox_server_version:  str|None        = None
        self.folders:                          CDispatch|None  = None
        self.offline:                          bool|None       = None
        self.session:                          CDispatch|None  = None
        self.stores:                           CDispatch|None  = None
        self.sync_objects:                     CDispatch|None  = None

        # Extract attributes
        attrs_map = {
            'Accounts':                      'accounts',
            'AddressLists':                  'address_lists',
            'AutoDiscoverConnectionMode':    'auto_discover_connection_mode',
            'AutoDiscoverXml':               'auto_discover_xml',
            'Categories':                    'categories',
            'CurrentProfileName':            'current_profile_name',
            'CurrentUser':                   'current_user',
            'DefaultStore':                  'default_store',
            'ExchangeConnectionMode':        'exchange_connection_mode',
            'ExchangeMailboxServerName':     'exchange_mailbox_server_name',
            'ExchangeMailboxServerVersion':  'exchange_mailbox_server_version',
            'Folders':                       'folders',
            'Offline':                       'offline',
            'Session':                       'session',
            'Stores':                        'stores',
            'SyncObjects':                   'sync_objects',
        }
        extract_attributes(self, self._namespace, attrs_map)
        return

    def add_store(self, store: CDispatch) -> None:
        '''
        Adds a Personal Folders (`.pst`) file to the current profile.

        Parameters
        ----------
        store : CDispatch
            The path of the `.pst` file to be added to the profile. If the
            `.pst` file does not exist, Microsoft Outlook creates it.
        
        Remarks
        -------
        Use the `remove_store` method to remove a `.pst` that is already added
        to a profile.
        '''
        return self._namespace.AddStore(store)
    
    def add_store_ex(self, store: CDispatch, type_: CDispatch) -> None:
        '''
        Adds a Personal Folders file (`.pst`) in the specified format to the
        current profile.

        Parameters
        ----------
        store : CDispatch
            The path of the `.pst` file to be added to the profile. If the
            `.pst` file does not exist, Microsoft Outlook creates it.
        type_: CDispatch
            The format in which the data file should be created.

        Returns
        -------
        None

        Remarks
        -------
        Use the `type_` constant to add a new `.pst` file that has greater
        storage capacity for items and folders and supports multilingual
        Unicode data, to the user's profile. The `olStoreANSI` constant allows
        you to create `.pst` files that do not provide full support for
        multilingual Unicode data, but are compatible with earlier versions of
        Outlook. The `type_` constant helps you create a `.pst` file in the
        default format that is compatible with the mailbox mode in which
        Outlook runs on the Microsoft Exchange Server.
        '''
        return self._namespace.AddStoreEx(store, type_)
    
    def compare_entry_ids(
            self,
            first_entry_id: str,
            second_entry_id: str,
    ) -> bool:
        '''
        Returns a boolean that indicates if two entry ID values refer to the
        same Outlook item.

        Parameters
        ----------
        first_entry_id : str
            The first entry ID to be compared.
        second_entry_id : str
        The second entry ID to be compared.

        Returns
        -------
        result : bool
            `True` if the entry ID values refer to the same Outlook item;
            otherwise, `False`.

        Remarks
        -------
        Entry identifiers cannot be compared directly because one object can be
        represented by two different binary values. Use this method to
        determine whether two entry identifiers represent the same object.
        '''
        return self._namespace.CompareEntryIDs(first_entry_id, second_entry_id)
    
    def create_contact_card(self, address_entry: CDispatch) -> CDispatch:
        '''
        Creates an instance of a `ContactCard` object for the contact that is
        specified by the `address_entry` parameter.

        Parameters
        ----------
        address_entry : CDispatch
            The `AddressEntry` object that represents the user for whom the
            contact card is to be created.

        Returns
        -------
        contact_card : CDispatch
            Returns a `ContactCard` object that is created for the specified
            user.

        Remarks
        -------
        The `ContactCard` object is available in the type library of Microsoft
        Office. Before calling `create_contact_card` to create a contact card
        in Microsoft Outlook, Outlook must be logged into an Outlook session.

        The `address_entry` parameter is an `AddressEntry` object that
        represents one of the following `AddressEntry` types defined in the
        `OlAddressEntryUserType` enumeration:

        - `olExchangeUserAddressEntry`
            - **Value:** 0
            - **Description:** An Exchange user that belongs to the same
              Exchange forest.
        - `olExchangeDistributionListAddressEntry`
            - **Value:** 1
            - **Description:** An address entry that is an Exchange
              distribution list.
        - `olExchangePublicFolderAddressEntry`
            - **Value:** 2
            - **Description:** An address entry that is an Exchange public
              folder.
        - `olExchangeAgentAddressEntry`
            - **Value:** 3
            - **Description:** An address entry that is an Exchange agent.
        - `olExchangeOrganizationAddressEntry`
            - **Value:** 4
            - **Description:** An address entry that is an Exchange
              organization.
        - `olExchangeRemoteUserAddressEntry`
            - **Value:** 5
            - **Description:** An Exchange user that belongs to a different
              Exchange forest.
        - `olOutlookContactAddressEntry`
            - **Value:** 10
            - **Description:** An address entry in an Outlook Contacts folder.
        - `olOutlookDistributionListAddressEntry`
            - **Value:** 11
            - **Description:** An address entry that is an Outlook distribution
              list.
        - `olLdapAddressEntry`
            - **Value:** 20
            - **Description:** An address entry that uses the Lightweight
              Directory Access Protocol (LDAP).
        - `olSmtpAddressEntry`
            - **Value:** 30
            - **Description:** An address entry that uses the Simple Mail
              Transfer Protocol (SMTP).
        - `olOtherAddressEntry`
            - **Value:** 40
            - **Description:** A custom or some other type of address entry
              such as FAX.

        Outlook raises the `E_INVALIDARG` error when you pass any of the
        following `OlAddressEntryUserType` values as an argument to the
        create_contact_card method:

        (missing)
        '''
        return self._namespace.CreateContactCard(address_entry)
    
    def create_recipient(self, recipient_name: str) -> CDispatch:
        '''
        Creates a `Recipient` object.

        Parameters
        ----------
        recipient_name : str
            The name of the recipient; it can be a string representing the
            display name, the alias, or the full SMTP e-mail address of the
            recipient.

        Returns
        -------
        recipient : CDispatch
            A `Recipient` object that represents the new recipient.

        Remarks
        -------
        This method is most commonly used to create a `Recipient` object for
        use with the `get_shared_default_folder` method, for example, to open a
        delegator's folder. It can also be used to verify a given name against
        an address book.
        '''
        return self._namespace.CreateRecipient(recipient_name)
    
    def create_sharing_item(
        self,
        context: str|CDispatch,
        provider: Optional[CDispatch]=None,
    ) -> CDispatch:
        '''
        Creates a new `SharingItem` object.

        Parameters
        ----------
        context : str | CDispatch
            Either a string value or a `Folder` object representing the sharing
            context to be used.
        provider : CDispatch | None, default: None
            An `OlSharingProvider` value representing the sharing provider to
            be used.

        Returns
        -------
        sharing_item : CDispatch
            A `SharingItem` object that represents a sharing message for the
            specified context.

        Remarks
        -------
        If a string value is specified in `context`, the method assumes that a
        URL has been provided as a sharing context. If a `Folder` object is
        specified in context, the method attempts to discover the sharing
        context from the folder. If no sharing context exists, or if more than
        one sharing context exists, an error occurs.

        If `provider` is not specified, the method attempts to use the
        appropriate sharing provider for the value specified in `context`.
        '''
        return self._namespace.CreateSharingItem(context, provider)
    
    def dial(self, contact_item: CDispatch) -> None:
        '''
        Displays the New Call dialog box that allows users to dial the primary
        phone number of a specified contact.

        Parameters
        ----------
        contact_item : CDispatch
            The `ContactItem` object of the contact you want to dial.

        Returns
        -------
        None
        '''
        return self._namespace.Dial(contact_item)
    
    def get_address_entry_from_id(self, id_: str) -> CDispatch:
        '''
        Returns an `AddressEntry` object that represents the address entry
        specified by ID.

        Parameters
        ----------
        id_ : str
            A string identifier for an address entry maintained for the
            session.

        Returns
        -------
        address_entry : CDispatch
            An `AddressEntry` that has the ID property matching the specified
            ID.

        Remarks
        -------
        The ID property for an `AddressEntry` is a permanent, unique string
        identifier that the transport provider assigns when an `AddressEntry`
        is created.

        Outlook maintains a hierarchy of address books for a session, and the
        address entry returned must match the given ID and be in one of the
        address books.

        `get_address_entry_from_id` returns an error if no item with the given
        ID can be found.

        `get_address_entry_from_id` also returns an error if no connection is
        available or the user is set to work offline.
        '''
        return self._namespace.GetAddressEntryFromID(id_)
    
    def get_default_folder(self, folder_type: int) -> CDispatch:
        '''
        Returns a `Folder` object that represents the default folder of the
        requested type for the current profile; for example, obtains the
        default `Calendar` folder for the user who is currently logged on.

        Parameters
        ----------
        folder_type : int
            The type of default folder to return.

        Returns
        -------
        mapi_folder : CDispatch
            A `Folder` object that represents the default folder of the
            requested type for the current profile.

        Remarks
        -------
        To return a specific non-default folder, use the `Folders` collection.

        If the default folder of the requested type does not exist, depending
        on the type, Outlook may create and return the folder, or may raise an
        error. For example, if `olFolderManagedEmail` is specified as the
        `folder_type` but the Managed Folders group has not been deployed,
        Microsoft Outlook raises an error.
        '''
        return self._namespace.GetDefaultFolder(folder_type)
    
    # def abc(self):
    #     '''
    #     description

    #     Parameters
    #     ----------

    #     Returns
    #     -------

    #     Remarks
    #     -------
        
    #     '''
    #     raise NotImplementedError('Not implemented')


class Inbox:
    '''
    The Outlook inbox folder.

    Attributes
    ----------
    address_book_name : str
        Returns or sets a String that indicates the Address Book
        name for the Folder object representing a Contacts folder. Read/write.
    application : CDispatch
        Returns an Application object that represents the parent Outlook
        application for the object. Read-only.
    class_ : int
        Returns an OlObjectClass constant indicating the object's class.
        Read-only.
    current_view : CDispatch
        Returns a View object representing the current view. Read-only.
    custom_views_only : bool
        Returns or sets a Boolean (bool in C#) that determines which views are
        displayed on the View menu for a given folder. Read/write.
    default_item_type : int
        Returns a constant from the OlItemType enumeration indicating the
        default Outlook item type contained in the folder. Read-only.
    default_message_class : str
        Returns a String representing the default message class
        for items in the folder. Read-only.
    description : str
        Returns or sets a String representing the description of
        the folder. Read/write.
    entry_id : str
        Returns a String representing the unique Entry ID of the
        object. Read-only.
    folder_path : str
        Returns a String that indicates the path of the current
        folder. Read-only.
    folders : CDispatch
        Returns the Folders collection that represents all the folders
        contained in the specified Folder. Read-only.
    in_app_folder_sync_object : bool
        Returns or sets a Boolean (bool in C#) that determines if the specified
        folder will be synchronized with the e-mail server. Read/write.
    is_sharepoint_folder : bool
        Returns a Boolean (bool in C#) that determines if the folder is a
        Microsoft SharePoint Foundation folder. Read-only.
    items : CDispatch
        Returns an Items collection object as a collection of Outlook items in
        the specified folder. Read-only.
    name : str
        Returns or sets a String value that represents the
        display name for the object. Read/write.
    parent : CDispatch
        Returns the parent Object of the specified object. Read-only.
    property_accessor : CDispatch
        Returns a PropertyAccessor object that supports creating, getting,
        setting, and deleting properties of the parent Folder object.
        Read-only.
    session : CDispatch
        Returns the NameSpace object for the current session. Read-only.
    show_as_outlook_ab : bool
        Returns or sets a Boolean (bool in C#) value that specifies whether
        the contact items folder will be displayed as an address list in the
        Outlook Address Book. Read/write.
    show_item_count : int
        Sets or returns a constant in the OlShowItemCount enumeration that
        indicates whether to display the number of unread messages in the
        folder or the total number of items in the folder in the Navigation
        Pane. Read/write.
    store : CDispatch
        Returns a Store object representing the store that contains the Folder
        object. Read-only.
    store_id : str
        Returns a String indicating the store ID for the folder.
        Read-only.
    unread_item_count : int
        Returns an Integer (int in C#) value indicating the number of unread
        items in the folder. Read-only.
    user_defined_properties : CDispatch
        Returns a UserDefinedProperties object that represents the user-defined
        custom properties for the Folder object. Read-only.
    views : CDispatch
        Returns the Views collection of the Folder object. Read-only.
    web_view_on : bool
        Returns or sets a Boolean (bool in C#) indicating the Web view state
        for a folder. Read/write.
    web_view_url : str
        Returns or sets a String indicating the URL of the Web
        page that is assigned to a folder. Read/write.
    '''
    
    def __init__(self) -> None:
        self.application = Application()
        self.namespace = self.application.get_namespace('MAPI')
        self._inbox = self.namespace._namespace.GetDefaultFolder(INBOX_FOLDER_NUMBER)

        # Type hints
        self.address_book_name:          str|None        = None
        self.class_:                     int|None        = None
        self.current_view:               CDispatch|None  = None
        self.custom_views_only:          bool|None       = None
        self.default_item_type:          int|None        = None
        self.default_message_class:      str|None        = None
        self.description:                str|None        = None
        self.entry_id:                   str|None        = None
        self.folder_path:                str|None        = None
        self.folders:                    CDispatch|None  = None
        self.in_app_folder_sync_object:  bool|None       = None
        self.is_sharepoint_folder:       bool|None       = None
        self.items:                      CDispatch|None  = None
        self.name:                       str|None        = None
        self.parent:                     CDispatch|None  = None
        self.property_accessor:          CDispatch|None  = None
        self.session:                    CDispatch|None  = None
        self.show_as_outlook_ab:         bool|None       = None
        self.show_item_count:            int|None        = None
        self.store:                      CDispatch|None  = None
        self.store_id:                   str|None        = None
        self.unread_item_count:          int|None        = None
        self.user_defined_properties:    CDispatch|None  = None
        self.views:                      CDispatch|None  = None
        self.web_view_on:                bool|None       = None
        self.web_view_url:               str|None        = None

        all_attrs = {
            'AddressBookName':         'address_book_name',
            'Class':                   'class_',
            'CurrentView':             'current_view',
            'DefaultItemType':         'default_item_type',
            'DefaultMessageClass':     'default_message_class',
            'Description':             'description',
            'EntryID':                 'entry_id',
            'FolderPath':              'folder_path',
            'Folders':                 'folders',
            'InAppFolderSyncObject':   'in_app_folder_sync_object',
            'IsSharePointFolder':      'is_sharepoint_folder',
            'Items':                   'items',
            'Name':                    'name',
            'Parent':                  'parent',
            'PropertyAccessor':        'property_accessor',
            'Session':                 'session',
            'ShowAsOutlookAB':         'show_as_outlook_ab',
            'ShowItemCount':           'show_item_count',
            'Store':                   'store',
            'StoreID':                 'store_id',
            'UnReadItemCount':         'unread_item_count',
            'UserDefinedProperties':   'user_defined_properties',
            'Views':                   'views',
            'WebViewOn':               'web_view_on',
            'WebViewURL':              'web_view_url',
        }
        for com_name, new_name in all_attrs.items():
            try:
                value = getattr(self._inbox, com_name)
                setattr(self, new_name, value)
                print(new_name.ljust(40), value)
            except:
                pass
        return


if __name__ == '__main__':
    error_code = main()
    exit(error_code)
