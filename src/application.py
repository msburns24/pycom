from __future__ import annotations
from typing import Optional, TYPE_CHECKING
from win32com.client import Dispatch, CDispatch
from .utils import extract_attributes
from .namespace import NameSpace


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