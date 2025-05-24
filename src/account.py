from __future__ import annotations
from typing import Optional, TYPE_CHECKING
from win32com.client import CDispatch
import _enums


class Account:
    '''
    The Account object represents an account defined for the current profile.

    Properties
    ----------
    account_type : OlAccountType
        Returns a constant in the `OlAccountType` enumeration that indicates
        the type of the Account. Read-only.
    auto_discover_connection_mode : OlAutoDiscoverConnectionMode
        Specifies the type of connection to the Exchange server for the
        auto-discovery service.
    auto_discover_xml : str
        Returns a string that represents information in XML retrieved from
        the auto-discovery service of the Microsoft Exchange Server that is
        associated with the account. Read-only.
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
    smtp_address : str
        Returns a string representing the Simple Mail Transfer Protocol (SMTP)
        address for the Account. Read-only.

    Methods
    -------
    '''
    
    def __init__(self, account: CDispatch) -> None:
        self._account = account
        return
    
    @property
    def account_type(self) -> _enums.OlAccountType:
        '''Returns a constant in the `OlAccountType` enumeration that indicates
        the type of the Account. Read-only.'''
        acct_type = self._account.AccountType
        return _enums.OlAccountType(acct_type)
    
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
    