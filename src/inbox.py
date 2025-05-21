from __future__ import annotations
from typing import Optional, TYPE_CHECKING
from win32com.client import Dispatch, CDispatch
from .utils import extract_attributes

if TYPE_CHECKING:
    from .application import Application


INBOX_FOLDER_NUMBER = 6


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