from __future__ import annotations
from typing import Optional, TYPE_CHECKING
from win32com.client import Dispatch, CDispatch
from .utils import extract_attributes

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

    def __init__(self, application: Application, namespace: CDispatch) -> None:
        self.application = application
        self._namespace = namespace
        return

    @property
    def accounts(self) -> CDispatch:
        '''An `Accounts` collection object that represents all the `Account`
        objects in the current profile. Read-only.'''
        return self._namespace.Accounts
    
    @property
    def address_lists(self) -> CDispatch:
        '''An `AddressLists` collection representing a collection of the
        address lists available for this session. Read-only.'''
        return self._namespace.AddressLists
    
    @property
    def auto_discover_connection_mode(self) -> int:
        '''A constant that specifies the type of connection to the Exchange
        server for auto-discovery service. Read-only.'''
        return self._namespace.AutoDiscoverConnectionMode
    
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
        return self._namespace.ExchangeConnectionMode
    
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
    
    def get_folder_from_id(
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
            An Object value that represents the specified Outlook item.

        Remarks
        -------
        This method is used for ease of transition between MAPI and
        OLE/Messaging applications and Outlook.
        '''
        return self._namespace.GetItemFromID(entry_id_item, entry_id_store)

    def get_global_address_list(self) -> CDispatch:
        '''
        Returns an `AddressList` object that represents the Exchange Global
        Address List.

        Parameters
        ----------
        None

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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    # def abc(self) -> CDispatch:
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

    