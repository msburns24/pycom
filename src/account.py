from __future__ import annotations
from typing import Optional, TYPE_CHECKING
from win32com.client import CDispatch
from . import _enums


if TYPE_CHECKING:
    from application import Application
    from namespace import NameSpace


class Account:
    '''
    The Account object represents an account defined for the current profile.

    Properties
    ----------
    account_type : OlAccountType
        Returns a constant in the `OlAccountType` enumeration that indicates
        the type of the Account. Read-only.
    application : Application
        Returns an `Application` object that represents the parent Outlook
        application for the object.
    auto_discover_connection_mode : OlAutoDiscoverConnectionMode
        Specifies the type of connection to the Exchange server for the
        auto-discovery service.
    auto_discover_xml : str
        Returns a string that represents information in XML retrieved from
        the auto-discovery service of the Microsoft Exchange Server that is
        associated with the account. Read-only.
    current_user : str
        Returns a string that represents the current user identity for the
        account. Read-only.
    delivery_store : str | None
        Returns a `Store` object that represents the default delivery store for
        the account. Returns `None` if the account does not have a default
        delivery store.
    display_name : str
        Returns a string representing the display name of the e-mail Account.
        Read-only.
    exchange_connection_mode : OlExchangeConnectionMode
        Returns an `OlExchangeConnectionMode` constant that indicates the
        current connection mode for the Microsoft Exchange Server that hosts
        the account mailbox. Read-only.
    exchange_mailbox_server_name : str
        Returns a string value that represents the name of the Microsoft
        Exchange Server that hosts the account mailbox. Read-only.
    exchange_mailbox_server_version : str
        Returns a string that represents the full version number of the
        Microsoft Exchange Server that hosts the account mailbox. Read-only.
    namespace : NameSpace
        Returns an `NameSpace` object that represents the parent namespace for
        the account.
    smtp_address : str
        Returns a string representing the Simple Mail Transfer Protocol (SMTP)
        address for the Account. Read-only.

    Methods
    -------
    get_address_entry_from_id(id_)
        Returns an `AddressEntry` object that represents the address entry
        specified by the given entry ID.
    get_recipient_from_id(entry_id)
        Returns the `Recipient` object that is identified by the given entry
        ID.

    Remarks
    -------
    This is a .NET interface derived from a COM co-class that is required by
    managed code for interoperability with the corresponding COM object. Use
    this derived interface to access all method, property, and event members of
    the COM object. However, if a method or event you want to use shares the
    same name under the same COM object, cast to the corresponding primary
    interface to call the method, and cast to the latest events interface to
    connect to the event. Refer to this topic for information about the COM
    object.
    '''
    
    def __init__(
            self,
            namespace: NameSpace,
            account: CDispatch
    ) -> None:
        self.namespace = namespace
        self._account = account
        return
    
    def __repr__(self) -> str:
        return f"<Account '{self.display_name}'>"
    
    @property
    def account_type(self) -> _enums.OlAccountType:
        '''Returns a constant in the `OlAccountType` enumeration that indicates
        the type of the Account. Read-only.'''
        acct_type = self._account.AccountType
        return _enums.OlAccountType(acct_type)
    
    @property
    def application(self) -> Application:
        '''Returns an `Application` object that represents the parent Outlook
        application for the object. Read-only.'''
        return self.namespace.application
    
    @property
    def auto_discover_connection_mode(
            self
    ) -> _enums.OlAutoDiscoverConnectionMode:
        '''Specifies the type of connection to the Exchange server for the
        auto-discovery service.'''
        conn_mode = self._account.AutoDiscoverConnectionMode
        return _enums.OlAutoDiscoverConnectionMode(conn_mode)
    
    @property
    def auto_discover_xml(self) -> str:
        '''
        Returns a string that represents information in XML retrieved from
        the auto-discovery service of the Microsoft Exchange Server that is
        associated with the account. Read-only.

        Remarks
        -------
        This property is similar to the `auto_discover_xml` property of the
        `NameSpace` object, except that this property applies to the account
        for which auto-discovery is completed and not necessarily to the
        primary Exchange account.

        The returned string of XML contains information about various Web
        services (for example, availability service and unified messaging
        service) and available servers.

        An error is returned if the account is not associated with an Exchange
        Server that is running Microsoft Exchange Server 2007 or later.
        '''
        return self._account.AutoDiscoverXml
    
    @property
    def current_user(self) -> str:
        '''
        Returns a string that represents the current user identity for the
        account. Read-only.
        '''
        return self._account.CurrentUser()
    
    @property
    def delivery_store(self) -> CDispatch | None:
        '''
        Returns a `Store` object that represents the default delivery store for
        the account. Returns `None` if the account does not have a default
        delivery store.
        '''
        return self._account.DeliveryStore
    
    @property
    def display_name(self) -> str:
        '''
        Returns a string representing the display name of the e-mail Account.
        Read-only.
        '''
        return self._account.DisplayName
    
    @property
    def exchange_connection_mode(self) -> _enums.OlExchangeConnectionMode:
        '''Returns an `OlExchangeConnectionMode` constant that indicates the
        current connection mode for the Microsoft Exchange Server that hosts
        the account mailbox. Read-only.'''
        conn_mode = self._account.ExchangeConnectionMode
        return _enums.OlExchangeConnectionMode(conn_mode)
    
    @property
    def exchange_mailbox_server_name(self) -> str:
        '''
        Returns a string value that represents the name of the Microsoft
        Exchange Server that hosts the account mailbox. Read-only.

        Remarks
        -------
        This property is similar to the `exchange_mailbox_server_name` property
        of the `NameSpace` object, except that this property applies to the
        Exchange Server that hosts the account mailbox, and not necessarily to
        the primary Exchange account.

        If an Exchange mailbox is not associated with this account, this
        property returns an empty string.
        '''
        return self._account.ExchangeMailboxServerName
    
    @property
    def exchange_mailbox_server_version(self) -> str:
        '''
        Returns a string that represents the full version number of the
        Microsoft Exchange Server that hosts the account mailbox. Read-only.

        Remarks
        -------
        This property is similar to the `exchange_mailbox_server_version`
        property of the `NameSpace` object, except that this property applies
        to the Exchange Server that hosts the account mailbox, and not
        necessarily to the primary Exchange account.

        This property returns a string that contains the version number of the
        Exchange server for the account. The version number has the following
        four parts:

        `<major version>.<minor version>.<build number>.<revision>`

        Not all parts may be present in the version number, depending on the
        version information that is supplied by the Exchange Server. For
        example, this property returns `"6.5.7638"` for Microsoft Exchange
        Server 2003 Service Pack 2.

        If an Exchange mailbox is not associated with this account, this
        property returns an empty string.
        '''
        return self._account.ExchangeMailboxServerVersion
    
    @property
    def iolk_account(self) -> CDispatch:
        raise NotImplementedError('Not implemented.')
    
    @property
    def smtp_address(self) -> str:
        '''
        Returns a string representing the Simple Mail Transfer Protocol (SMTP)
        address for the Account. Read-only.
        
        Remarks
        -------
        The purpose of `smtp_address` and `user_name` is to provide an
        account-based context to determine identity.

        If the account does not have an SMTP address, `smtp_address` returns an
        empty string.
        '''
        return self._account.SmtpAddress
    
    @property
    def user_name(self) -> str:
        '''
        Returns a string representing the user name for the Account. Read-only.

        Remarks
        -------
        The purpose of `smtp_address` and `user_name` is to provide an
        account-based context to determine identity.

        If the account does not have a user name defined, `user_name` returns
        an empty string.
        '''
        return self._account.UserName

    def get_address_entry_from_id(self, id_: str) -> CDispatch:
        '''
        Returns an `AddressEntry` object that represents the address entry
        specified by the given entry ID.

        Parameters
        ----------
        id_ : str
            Used to identify an address entry that is maintained for the
            session.

        Returns
        -------
        CDispatch
            An `AddressEntry` that has the ID property that matches the
            specified ID.

        Remarks
        -------
        This method is similar to the `get_address_entry_from_id` method of the
        `NameSpace` object, but has some additional contextual information
        about which account to use for the look-up. If there are multiple
        Microsoft Exchange accounts in the current profile, use the
        `get_address_entry_from_id` method for the corresponding account.

        The ID property for an `AddressEntry` is a permanent, unique string
        identifier that the transport provider assigns when an `AddressEntry`
        is created. Outlook maintains a hierarchy of address books for a
        session, and the address entry that is returned must match the given
        ID and be in one of the address books.

        `get_address_entry_from_id` returns an error if no item with the given
        ID can be found, if no connection is available, or if the user is set
        to work offline.
        '''
        return self._account.GetAddressEntryFromID(id_)

    def get_recipient_from_id(self, entry_id: str) -> CDispatch:
        '''
        Returns the `Recipient` object that is identified by the given entry
        ID.

        Parameters
        ----------
        entry_id : str
            The `entry_id` of the recipient.

        Returns
        -------
        CDispatch
            A `Recipient` object that represents the recipient associated with
            the specified entry ID.

        Remarks
        -------
        This method is similar to the `get_recipient_from_id` method of the
        `NameSpace` object. If there are multiple Microsoft Exchange accounts
        in the current profile, use the `get_recipient_from_id` method for the
        corresponding account.
        '''
        return self._account.GetRecipientFromID(entry_id)
