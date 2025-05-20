from rich import print
from win32com import client
from win32com.client import Dispatch, CDispatch


# Constants
INBOX_FOLDER_NUMBER = 6


def main() -> int:
    # app = Application()
    inbox = Inbox()
    return 0


class Application:
    '''
    Represents the entire Outlook application.

    Documentation taken from the `Microsoft.Office.Interop.Outlook.Application`
    [Documentation](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.application).

    Attributes
    ----------
    active_explorer : CDispatch
        Returns the topmost `Explorer` object on the desktop.
    active_window : CDispatch
        Returns an object representing the topmost Microsoft Outlook window on
        the desktop, either an `Explorer` or an `Inspector` object.
    assistance : CDispatch
        Returns an `IAssistance`
    class_ : int
        Returns an `OlObjectClass` constant indicating the object's class.
        Read-only.
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
    advanced_search(scope, filter, search_sub_folders, tag) -> CDispatch
        Performs a search based on a specified DAV Searching and Locating
        (DASL) search string.
    copy_file(file_path: str, dest_folder_path: str) -> CDispatch
        Copies a file from a specified location into a Micro1soft Outlook
        store.
    create_item(self, item_type: CDispatch) -> CDispatch
        Creates and returns a new Microsoft Outlook item.
    
    
    CreateItemFromTemplate
    CreateObject
    GetIDsOfNames
    GetObjectReference
    GetTypeInfo
    GetTypeInfoCount
    Invoke
    IsSearchSynchronous
    QueryInterface
    Quit
    RefreshFormRegionDefinition
    Release


    create_item_from_template()
        Creates a new Microsoft Outlook item from an Outlook template (`.oft`)
        and returns the new item.
    create_object()
        Creates an Automation object of the specified class.
    get_namespace(type_: str='MAPI') -> CDispatch
        Returns a NameSpace object of the specified type.
    get_object_reference()
        Creates a strong or weak object reference for a specified `Outlook`
        object.
    is_search_synchronous() -> bool
        Returns a boolean indicating if a search will be synchronous or
        asynchronous.
    refresh_form_region_definition()
        Refreshes the cache by obtaining the current definition from the
        Windows registry for one or all of the form regions that are defined
        for the local machine and the current user.
    '''

    def __init__(self) -> None:
        self._application = Dispatch('Outlook.Application')

        self.assistance: CDispatch|None = None
        self.class_: int|None = None
        self.com_add_ins: CDispatch|None = None
        self.data_privacy_options: CDispatch|None = None
        self.default_profile_name: str|None = None
        self.explorers: CDispatch|None = None
        self.inspectors: CDispatch|None = None
        self.is_trusted: bool|None = None
        self.language_settings: CDispatch|None = None
        self.name: str|None = None
        self.picker_dialog: CDispatch|None = None
        self.product_code: str|None = None
        self.reminders: CDispatch|None = None
        self.session: CDispatch|None = None
        self.time_zones: CDispatch|None = None
        self.version: str|None = None

        all_attrs = {
            'Assistance':          'assistance',
            'Class':               'class_',
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
        for com_name, new_name in all_attrs.items():
            try:
                value = getattr(self._application, com_name)
                setattr(self, new_name, value)
            except:
                pass
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
            An Object value that represents the copied file.
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

    def create_item_from_template(self):
        '''
        Creates a new Microsoft Outlook item from an Outlook template (`.oft`)
        and returns the new item.
        '''
        raise NotImplementedError('Not implemented.')
    
    def create_object(self):
        '''
        Creates an Automation object of the specified class.
        '''
        raise NotImplementedError('Not implemented.')
    
    def get_namespace(self, type_: str='MAPI') -> CDispatch:
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
        return self._application.GetNamespace(type_)
    
    def get_object_reference(self):
        '''Creates a strong or weak object reference for a specified `Outlook`
        object.'''
        raise NotImplementedError('Not implemented.')
    
    def is_search_synchronous(self) -> bool:
        '''
        Returns a boolean indicating if a search will be synchronous or
        asynchronous.
        '''
        raise NotImplementedError('Not implemented.')
    
    def refresh_form_region_definition(self):
        '''
        Refreshes the cache by obtaining the current definition from the
        Windows registry for one or all of the form regions that are defined
        for the local machine and the current user.
        '''
        raise NotImplementedError('Not implemented.')
    


class Inbox:
    '''
    The Outlook inbox folder.

    Documentation taken from the `Microsoft.Office.Interop.Outlook.Folder`
    [Documentation](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.folder).

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
        application = Application()
        outlook = application.get_namespace('MAPI')
        self._inbox = outlook.GetDefaultFolder(INBOX_FOLDER_NUMBER)

        self.address_book_name: str | None = None
        self.application: CDispatch | None = None
        self.class_: int | None = None
        self.current_view: CDispatch | None = None
        self.custom_views_only: bool | None = None
        self.default_item_type: int | None = None
        self.default_message_class: str | None = None
        self.description: str | None = None
        self.entry_id: str | None = None
        self.folder_path: str | None = None
        self.folders: CDispatch | None = None
        self.in_app_folder_sync_object: bool | None = None
        self.is_sharepoint_folder: bool | None = None
        self.items: CDispatch | None = None
        self.name: str | None = None
        self.parent: CDispatch | None = None
        self.property_accessor: CDispatch | None = None
        self.session: CDispatch | None = None
        self.show_as_outlook_ab: bool | None = None
        self.show_item_count: int | None = None
        self.store: CDispatch | None = None
        self.store_id: str | None = None
        self.unread_item_count: int | None = None
        self.user_defined_properties: CDispatch | None = None
        self.views: CDispatch | None = None
        self.web_view_on: bool | None = None
        self.web_view_url: str | None = None

        all_attrs = {
            'AddressBookName':         'address_book_name',
            'Application':             'application',
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
