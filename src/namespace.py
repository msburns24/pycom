from __future__ import annotations
from typing import TYPE_CHECKING
from win32com.client import CDispatch
from . import _enums
from .account import Account

if TYPE_CHECKING:
    from .application import Application


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
    get_folder_from_id(entry_id_item, entry_id_store)
        Returns a Microsoft Outlook item identified by the specified entry ID
        (if valid).
    get_global_address_list()
        Returns an `AddressList` object that represents the Exchange Global
        Address List.
    get_ids_of_names
        Not implemented
    get_item_from_id(entry_id_item, entry_id_store)
        Returns a Microsoft Outlook item identified by the specified entry ID
        (if valid).
    get_recipient_from_id(entry_id)
        Returns a `Recipient` object identified by the specified entry ID
        (if valid).
    get_select_names_dialog()
        Obtains a `SelectNamesDialog` object for the current session.
    get_shared_default_folder(recipient, folder_type)
        Returns a `Folder` object that represents the specified default folder
        for the specified user.
    get_store_from_id(id_)
        Returns a `Store` object that represents the store specified by ID.
    logoff()
        Logs the user off from the current MAPI session.
    logon(profile, password, show_dialog, new_session)
        Logs the user on to MAPI, obtaining a MAPI session.
    open_shared_folder(path, name, download_attachments, use_ttl)
        Opens a shared folder referenced through a URL or file name.
    open_shared_item(path)
        Opens a shared item from a specified path or URL.
    pick_folder()
        Displays the Pick Folder dialog box.
    remove_store(folder)
        Removes a Personal Folders file (`.pst`) from the current MAPI profile
        or session.
    send_and_receive(show_progress_dialog)
        Initiates immediate delivery of all undelivered messages submitted in
        the current session, and immediate receipt of mail for all accounts in
        the current profile.
    '''

    def __init__(
            self,
            application: Application,
            namespace_type: str,
            namespace: CDispatch
    ) -> None:
        self.application = application
        self._namespace_type = namespace_type
        self._namespace = namespace
        return
    
    def __repr__(self) -> str:
        return f"<NameSpace '{self._namespace_type}'>"

    @property
    def accounts(self) -> list[Account]:
        '''An `Accounts` collection object that represents all the `Account`
        objects in the current profile. Read-only.'''
        accounts = self._namespace.Accounts
        return [Account(self, acct) for acct in accounts]
    
    @property
    def address_lists(self) -> CDispatch:
        '''An `AddressLists` collection representing a collection of the
        address lists available for this session. Read-only.'''
        return self._namespace.AddressLists
    
    @property
    def auto_discover_connection_mode(self) -> int:
        '''A constant that specifies the type of connection to the Exchange
        server for auto-discovery service. Read-only.'''
        auto_disc_conn_mode = self._namespace.AutoDiscoverConnectionMode
        return _enums.OlAutoDiscoverConnectionMode(auto_disc_conn_mode)
    
    @property
    def auto_discover_xml(self) -> str:
        '''Information in XML retrieved from the auto-discovery service of an
        Exchange server. Read-only.'''
        return self._namespace.AutoDiscoverXml
    
    @property
    def categories(self) -> CDispatch:
        '''The set of `Category` objects available to the namespace.
        Read/write.'''
        return self._namespace.Categories
    
    @property
    def current_profile_name(self) -> str:
        '''The name of the current profile. Read-only.'''
        return self._namespace.CurrentProfileName
    
    @property
    def current_user(self) -> CDispatch:
        '''The display name of the currently logged-on user as a `Recipient`
        object. Read-only.'''
        return self._namespace.CurrentUser
    
    @property
    def default_store(self) -> CDispatch:
        '''The default `Store` for the profile. Read-only.'''
        return self._namespace.DefaultStore
    
    @property
    def exchange_connection_mode(self) -> int:
        '''A constant that indicates the current connection mode the user is
        using. Read-only.'''
        conn_mode = self._namespace.ExchangeConnectionMode
        return _enums.OlExchangeConnectionMode(conn_mode)
    
    @property
    def exchange_mailbox_server_name(self) -> str:
        '''The name of the Exchange server on which the active mailbox is
        hosted. Read-only.'''
        return self._namespace.ExchangeMailboxServerName
    
    @property
    def exchange_mailbox_server_version(self) -> str:
        '''The full version of the Exchange server on which the active mailbox
        is hosted. Read-only.'''
        return self._namespace.ExchangeMailboxServerVersion
    
    @property
    def folders(self) -> CDispatch:
        '''All the folders contained in the specified NameSpace. Read-only.'''
        return self._namespace.Folders
    
    @property
    def offline(self) -> bool:
        '''Indicates `True` if Outlook is offline (not connected to an Exchange
        server), and `False` if online (connected to an Exchange server).
        Read-only.'''
        return self._namespace.Offline
    
    @property
    def session(self) -> CDispatch:
        '''The `NameSpace` object for the current session. Read-only.'''
        return self._namespace.Session
    
    @property
    def stores(self) -> CDispatch:
        '''All the `Store` objects in the current profile. Read-only.'''
        return self._namespace.Stores
    
    @property
    def sync_objects(self) -> CDispatch:
        '''Contains all Send/Receive groups. Read-only.'''
        return self._namespace.SyncObjects

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
        bool
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

        - `EXCHANGE_USER: 0` - An Exchange user that belongs to the same
          Exchange forest.
        - `EXCHANGE_DISTRIBUTION_LIST: 1` - An address entry that is an
          Exchange distribution list.
        - `EXCHANGE_PUBLIC_FOLDER: 2` - An address entry that is an Exchange
          public folder.
        - `EXCHANGE_AGENT: 3` - An address entry that is an Exchange agent.
        - `EXCHANGE_ORGANIZATION: 4` - An address entry that is an Exchange
          organization.
        - `EXCHANGE_REMOTE_USER: 5` - An Exchange user that belongs to a
          different Exchange forest.
        - `OUTLOOK_CONTACT: 10` -  An address entry in an Outlook Contacts
          folder.
        - `OUTLOOK_DISTRIBUTION_LIST: 11` -  An address entry that is an
          Outlook distribution list.
        - `LDAP: 20` -  An address entry that uses the Lightweight Directory
          Access Protocol (LDAP).
        - `SMTP: 30` -  An address entry that uses the Simple Mail Transfer
          Protocol (SMTP).
        - `OTHER: 40` -  A custom or some other type of address entry such as
          FAX.

        Outlook raises the `E_INVALIDARG` error when you pass any of the
        following `OlAddressEntryUserType` values as an argument to the
        create_contact_card method:

        (missing Microsoft docs)
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
        provider: int | _enums.OlSharingProvider,
    ) -> CDispatch:
        '''
        Creates a new `SharingItem` object.

        Parameters
        ----------
        context : str | CDispatch
            Either a string value or a `Folder` object representing the sharing
            context to be used.
        provider : int | OlSharingProvider
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
    
    def get_default_folder(
            self,
            folder_type: int | _enums.OlDefaultFolders
    ) -> CDispatch:
        '''
        Returns a `Folder` object that represents the default folder of the
        requested type for the current profile; for example, obtains the
        default `Calendar` folder for the user who is currently logged on.

        Parameters
        ----------
        folder_type : int | OlDefaultFolders
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
    
    def get_folder_from_id(
            self,
            entry_id_folder: str,
            entry_id_store: CDispatch
    ) -> CDispatch:
        '''
        Returns a Folder object identified by the specified entry ID
        (if valid).

        Parameters
        ----------
        entry_id_folder : str
            The `entry_id` of the folder.
        entry_id_store : CDispatch
            The `store_id` for the folder. `entry_id_store` usually must be
            provided when retrieving an item based on its MAPI IDs.

        Returns
        -------
        obj : CDispatch
            An Object value that represents the specified Outlook item.

        Remarks
        -------
        This method is used for ease of transition between MAPI and
        OLE/Messaging applications and Outlook.
        '''
        return self._namespace.GetItemFromID(entry_id_folder, entry_id_store)

    def get_global_address_list(self) -> CDispatch:
        '''
        Returns an `AddressList` object that represents the Exchange Global
        Address List.

        Returns
        -------
        address_list : CDispatch
            An `AddressList` that represents the Global Address List.

        Remarks
        -------
        `get_global_address_list` supports only Exchange servers. It returns an
        error if the Global Address List is not available or cannot be found.

        It also returns an error if no connection is available or the user is
        set to work offline.
        '''
        return self._namespace.GetGlobalAddressList()

    def get_ids_of_names(self) -> CDispatch:
        # return self._namespace.GetIDsOfNames()
        raise NotImplementedError('Not implemented')

    def get_item_from_id(
            self,
            entry_id_item: str,
            entry_id_store: CDispatch
    ) -> CDispatch:
        '''
        Returns a Microsoft Outlook item identified by the specified entry ID
        (if valid).

        Parameters
        ----------
        entry_id_item : str
            The `entry_id` of the item.
        entry_id_store : CDispatch
            The `store_id` for the folder. `entry_id_store` usually must be
            provided when retrieving an item based on its MAPI IDs.

        Returns
        -------
        obj : CDispatch
            An `Object` value that represents the specified Outlook item.

        Remarks
        -------
        This method is used for ease of transition between MAPI and
        OLE/Messaging applications and Outlook.
        '''
        return self._namespace.GetItemFromID(entry_id_item, entry_id_store)

    def get_recipient_from_id(self, entry_id: str) -> CDispatch:
        '''
        Returns a `Recipient` object identified by the specified entry ID
        (if valid).

        Parameters
        ----------
        entry_id : str
            The `entry_id` of the recipient.

        Returns
        -------
        recipient : CDispatch
            A `Recipient` object that represents the specified recipient.

        Remarks
        -------
        This method is used for ease of transition between MAPI and
        OLE/Messaging applications and Microsoft Outlook.
        '''
        return self._namespace.GetRecipientFromID(entry_id)

    def get_select_names_dialog(self) -> CDispatch:
        '''
        Obtains a `SelectNamesDialog` object for the current session.

        Returns
        -------
        select_names_dialog : CDispatch
            A `SelectNamesDialog` object for the current session. The
            `SelectNamesDialog` object supports displaying the Select Names
            dialog box for the user to select entries from one or more address
            lists in the current session.
        '''
        return self._namespace.GetSelectNamesDialog()

    def get_shared_default_folder(
            self,
            recipient: CDispatch,
            folder_type: int | _enums.OlDefaultFolders
    ) -> CDispatch:
        '''
        Returns a `Folder` object that represents the specified default folder
        for the specified user.

        Parameters
        ----------
        recipient : CDispatch 
            The owner of the folder. Note that the `recipient` object must be
            resolved.
        folder_type : int | OlDefaultFolders 
            The type of folder.

        Returns
        -------
        mapi_folder : CDispatch
            A `Folder` object that represents the specified default folder for
            the specified user.

        Remarks
        -------
        This method is used in a delegation scenario, where one user has
        delegated access to another user for one or more of their default
        folders (for example, their shared Calendar folder).
        '''
        return self._namespace.GetSharedDefaultFolder(recipient, folder_type)

    def get_store_from_id(self, id_: str) -> CDispatch:
        '''
        Returns a `Store` object that represents the store specified by ID.

        Parameters
        ----------
        id_ : str
            A string value identifying a store.

        Returns
        -------
        store : CDispatch
            A `Store` object that has the `store_id` property matching ID.

        Remarks
        -------
        The `store_id` property of a Store is unique to the profile for the
        session. It is equivalent to the `MAPI` property
        `pid_tag_store_entry_id`.

        The store must be mounted in order for this method to succeed.

        `get_store_from_id` returns an error if no store with the specified ID
        can be found for the current session.
        '''
        return self._namespace.GetStoreFromID(id_)

    def logoff(self) -> None:
        '''
        Logs the user off from the current MAPI session.
        '''
        return self._namespace.Logoff()

    def logon(
            self,
            profile: str,
            password : str,
            show_dialog: bool,
            new_session: bool,
    ) -> None:
        '''
        Logs the user on to MAPI, obtaining a MAPI session.

        Parameters
        ----------
        profile : str
            The MAPI profile name, as a string, to use for the session.
            Specify an empty string to use the default profile for the current
            session.
        password : str
            The password (if any), as a string, associated with the profile.
            This parameter exists only for backwards compatibility and for
            security reasons, it is not recommended for use. Microsoft Outlook
            will prompt the user to specify a password in most system
            configurations. This is your logon password and should not be
            confused with PST passwords.
        show_dialog : bool
            `True` to display the MAPI logon dialog box to allow the user to
            select a MAPI profile.
        new_session : bool
            `True` to create a new Outlook session. Since multiple sessions
            cannot be created in Outlook, this parameter should be specified
            as `True` only if a session does not already exist.

        Remarks
        -------
        Use the `logon` method only to log on to a specific profile when
        Outlook is not already running. This is because only one Outlook
        process can run at a time, and that Outlook process uses only one
        profile and supports only one MAPI session. When a user start Outlook a
        second time, that instance of Outlook runs within the same Outlook
        process, does not create a new process, and uses the same profile.

        If Outlook is already running, using this method does not create a new
        Outlook session or change the current profile to a different one.
        '''
        return self._namespace.Logon(profile,
                                     password,
                                     show_dialog,
                                     new_session)

    def open_shared_folder(
            self,
            path: str,
            name: str,
            download_attachments: bool,
            use_ttl: bool,
    ) -> CDispatch:
        '''
        Opens a shared folder referenced through a URL or file name.

        Parameters
        ----------
        path : str
            The URL or local file name of the shared folder to be opened.
        name : str
            The name of the Really Simple Syndication (RSS) feed or Webcal
            calendar. This parameter is ignored for other shared folder types.
        download_attachments : bool
            Indicates whether to download enclosures (for RSS feeds) or
            attachments (for Webcal calendars.) This parameter is ignored for
            other shared folder types.
        use_ttl : bool
            Indicates whether the Time To Live (TTL) setting in an RSS feed or
            WebCal calendar should be used. This parameter is ignored for other
            shared folder types.

        Returns
        -------
        mapi_folder : CDispatch
            A `Folder` object that represents the shared folder.

        Remarks
        -------
        This method does not support iCalendar appointment (`.ics`) files. To
        open iCalendar appointment files, you can use the
        `open_shared_item` method of the `NameSpace` object.

        You can use the `get_shared_default_folder` method of the `NameSpace`
        object to share default folders, such as the Inbox folder, in Exchange.
        '''
        return self._namespace.OpenSharedFolder(path,
                                                name,
                                                download_attachments,
                                                use_ttl)

    def open_shared_item(self, path: str) -> CDispatch:
        '''
        Opens a shared item from a specified path or URL.

        Parameters
        ----------
        path : str
            The path or URL of the shared item to be opened.

        Returns
        -------
        CDispatch
            An `Object` representing the appropriate Outlook item for the
            shared item.

        Remarks
        -------
        This method is used to open iCalendar appointment (`.ics`) files, vCard
        (`.vcf`) files, and Outlook message (`.msg`) files. The type of object
        returned by this method depends on the type of shared item opened.
        '''
        return self._namespace.OpenSharedItem(path)

    def pick_folder(self) -> CDispatch:
        '''
        Displays the Pick Folder dialog box.

        Returns
        -------
        CDispatch
            A `Folder` object that represents the folder that the user selects
            in the dialog box, or Nothing if the dialog box is canceled by the
            user.

        Remarks
        -------
        The Pick Folder dialog box is a modal dialog box which means that code
        execution will not continue until the user either selects a folder or
        cancels the dialog box.
        '''
        return self._namespace.PickFolder()

    def remove_store(self, folder: CDispatch) -> None:
        '''
        Removes a Personal Folders file (`.pst`) from the current MAPI profile
        or session.

        Parameters
        ----------
        folder : CDispatch
            The Personal Folders file (`.pst`) to be deleted from the list.

        Remarks
        -------
        This method removes a store only from the Microsoft Outlook user
        interface. You cannot remove a store from the main mailbox on the
        server or from a user's hard disk using the Outlook object model.
        '''
        return self._namespace.RemoveStore(folder)

    def send_and_receive(
            self,
            show_progress_dialog: bool=True
    ) -> None:
        '''
        Initiates immediate delivery of all undelivered messages submitted in
        the current session, and immediate receipt of mail for all accounts in
        the current profile.

        Parameters
        ----------
        show_progress_dialog : bool, default: True
            Indicates whether the Outlook Send/Receive Progress dialog box
            should be displayed, regardless of user settings.

        Remarks
        -------
        Calling the `send_and_receive` method is asynchronous.

        `send_and_receive` provides the programmatic equivalent to the
        Send/Receive All command that is available when you click Tools and
        then Send/Receive.

        If you do not need to synchronize all objects, you can use the
        SyncObjects collection object to select specific objects. 

        All accounts defined in the current profile are used in Send/Receive
        All. If an online connection is required to perform the Send/Receive
        All, the connection is made according to user preferences.
        '''
        return self._namespace.SendAndReceive(show_progress_dialog)
