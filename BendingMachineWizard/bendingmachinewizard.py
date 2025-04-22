"""
Bendingmachine wizard allows the user to export BVBS data into attributes in Allplan on the object.
IFC export can be enabled upon request.
Missing Features
- Support for Variable placements (awaiting Allplan development)
- Support for spiral reinforcement (will not do)
- Support for bended circular shapes with more then one circular segment (will not do)
- Removing attributes upon creation of new ones (awaiting Allplan development)
- Meshes and other special placement shapes apart from BF2D and BF3D (will not do)

v0.2 - WIP created on 10/24 by Bert Van Overmeir for EDF.
"""

from inspect import Attribute
from pydoc import doc
from typing import Any, List, TYPE_CHECKING, cast
from enum import Enum
from pathlib import Path
import numpy as np
import os
import datetime
import re
import math

import NemAll_Python_IFW_ElementAdapter as AllplanElementAdapter
import NemAll_Python_IFW_Input as AllplanIFW
import NemAll_Python_BaseElements as AllplanBaseElements
import NemAll_Python_Utility as AllplanUtil
import NemAll_Python_AllplanSettings as AllplanSettings
import NemAll_Python_Reinforcement as AllplanReinforcement
from ServiceExamples import AttributeService
import Utils as Utils
import BuildingElementStringTable as BuildingElementStringTable
import AnyValueByType as AnyValueByType

from BuildingElement import BuildingElement
from BuildingElementComposite import BuildingElementComposite
from BuildingElementPaletteService import BuildingElementPaletteService
from StringTableService import StringTableService
from ControlProperties import ControlProperties
from BuildingElementListService import BuildingElementListService
from CreateElementResult import CreateElementResult
from BuildingElementTupleUtil import BuildingElementTupleUtil
import Utils.LibraryBitmapPreview
from BuildingElementAttributeList import BuildingElementAttributeList
from ControlPropertiesUtil import ControlPropertiesUtil


def create_preview(_build_ele: BuildingElement,
                   _doc      : AllplanElementAdapter.DocumentAdapter) -> CreateElementResult:
    """ Creation of the element preview

    Args:
        _build_ele: building element with the parameter properties
        _doc:       document of the Allplan drawing files

    Returns:
        created elements for the preview
    """

    return CreateElementResult(Utils.LibraryBitmapPreview.create_library_bitmap_preview(r"C:\Users\bovermeir\Documents\Nemetschek\Allplan\2025\Usr\Local\Library\BendingMachineWizard\bending.png"))

def check_allplan_version(build_ele, version):
    return True

def create_interactor(coord_input:               AllplanIFW.CoordinateInput,
                      pyp_path:                  str,
                      _global_str_table_service: StringTableService,
                      build_ele_list:            List[BuildingElement],
                      build_ele_composite:       BuildingElementComposite,
                      control_props_list:        List[ControlProperties],
                      modify_uuid_list:          list):
    interactor = BendingMachineWizardInteractor(coord_input, pyp_path, build_ele_list, build_ele_composite,
                                         control_props_list, modify_uuid_list)
    return interactor


class Event(Enum):
    NO_EVENT = 0
    USER_START_EXPORT = 1
    USER_CONFIRM_EXPORT = 2


class EventOrigin(Enum):
    BUTTONCLICK = 0
    SELECTIONCOMPLETE_SINGLE = 1
    SELECTIONCOMPLETE_MULTI = 2
    SELECTIONCOMPLETE_POINT = 3
    OTHER = 4


class ShapeType(Enum):
    SHAPE2D = 0
    SHAPE3D = 1


class BMWizardInfo(Enum):
    ERR_CREATING_NEW_ATTRIBUTES = 0
    ERR_ATTRIBUTES_ASSIGNMENT_FAILED = 1
    ERR_INVALID_IFC_TYPE_OR_MISSING = 2
    ERR_GENERAL_PARSING_ERROR = 3
    ERR_MATCHING_ALLPLAN_DATA = 4
    ERR_SELECTION_FAILED = 5
    ERR_ATTRIBUTES_UNDEFINED_IN_UI = 6
    ERR_GENERAL_EXPORT_BVBS_ERROR = 10
    ERR_GENERAL_IMPORT_BVBS_ERROR = 11
    ERR_COUPLER_MATCHING = 12
    INFO_IDLE = 13
    ERR_NOT_IMPLEMENTED = 15
    ERR_IFC_PATH_INVALID = 16
    ERR_IFC_EXPORT_FAILED = 17
    INFO_EXPORT_IFC = 18
    INFO_FINISHED = 19
    INFO_PREPARING_DATA = 20


class SelectionType(Enum):
    NONE = 0
    SINGLE_SELECTION = 1
    MULTISELECTION = 2
    FACE_SELECTION = 3
    INPUT_POINT = 4


class BendingMachineWizardInteractor():
    """
    Definition of class BendingMachineWizardInteractor
    """
    def __init__(self,
                 coord_input:           AllplanIFW.CoordinateInput,
                 pyp_path:              str,
                 build_ele_list:        List[BuildingElement],
                 build_ele_composite:   BuildingElementComposite,
                 control_props_list:    List[ControlProperties],
                 modify_uuid_list:      list):
        """
        Create the interactor

        Args:
            coord_input:               coordinate input
            pyp_path:                  path of the pyp file
            build_ele_list:            building element list
            build_ele_composite:       building element composite
            control_props_list:        control properties list
            modify_uuid_list:          UUIDs of the existing elements in the modification mode
        """

        self.coord_input         = coord_input
        self.pyp_path            = pyp_path
        self.build_ele_list      = build_ele_list
        self.build_ele_composite = build_ele_composite
        self.control_props_list  = control_props_list
        self.modify_uuid_list    = modify_uuid_list
        self.palette_service     = None
        self.model_ele_list      = []
        self.modification        = False
        self.close_interactor    = False
        self.user_origin_event   = Event.NO_EVENT
        self.user_selection_mode = SelectionType.NONE
        self.user_selection      = AllplanIFW.PostElementSelection()
        self.user_mulitselection_list = None
        self.user_referencepoints= []
        self.user_single_selection_list     = AllplanElementAdapter.BaseElementAdapter()
        self.user_filter         = None
        self.user_message        = ""
        self.is_second_input_point = False
        self.ctrl_prop_util      = None
        # start palette VIS
        self.palette_service = BuildingElementPaletteService(self.build_ele_list, self.build_ele_composite,
                                                             self.build_ele_list[0].script_name,
                                                             self.control_props_list, pyp_path + "\\", None)
        self.palette_service.show_palette(self.build_ele_list[0].pyp_file_name)
        (local_str_table, global_str_table) = self.build_ele_list[0].get_string_tables()
        self.ctrl_prop_util = ControlPropertiesUtil(control_props_list, build_ele_list)
        # workaround "list may not be empty upon visibility change" BUG
        temp_list = self.build_ele_list[0].AnyValueByTypeList.value
        temp_list.append(AnyValueByType.AnyValueByType("Text", " ", ""))
        # set startup vis
        self.set_tab_status_startup()
        AllplanHelpers.static_init(self.coord_input, local_str_table)
        AllplanHelpers.show_message_in_taskbar(AllplanHelpers.get_message(BMWizardInfo.INFO_IDLE))
        # init variables for events
        self.attribute_settings = None
        self.selected_elements = None
        self.assembly_match_table = None
        self.created_rebar = None

    def on_control_event(self, event_id):
        """ control the different ID's that can be called via buttons.
        """
        self.palette_service.on_control_event(event_id)
        self.palette_service.update_palette(-1, True)
        self.set_event(Event(event_id))
        ok = self.event_do(Event(event_id), EventOrigin.BUTTONCLICK)
        AllplanHelpers.show_message_in_taskbar(AllplanHelpers.get_message(BMWizardInfo.INFO_IDLE))
        if not ok:
            self.ctrl_prop_util.set_enable_function("yesbutton", self.disable_variable_function)
            self.build_ele_list[0].text_info_user.value = "Error"
            AllplanHelpers.finite_progressbar_stop()

    def disable_variable_function(self) -> bool:
        return False

    def enable_variable_function(self) -> bool:
        return True

    def process_mouse_msg(self, mouse_msg, pnt, msg_info):
        """
        Process user input depending on the defined selection mode by the program (SelectionType.xxx).<br>
        After the action is completed, the user will be directed towards the defined user_origin_event with flag EventOrigin.[SelectionType].<br>
        User input data is saved in user_single_selection_list, user_multiselection_list or user_referencepoints depending on SelectionType.
        """
        if self.get_selection_mode() == SelectionType.SINGLE_SELECTION:
            is_element_found = self.coord_input.SelectElement(mouse_msg,pnt,msg_info,True,True,True)
            if is_element_found:
                self.user_single_selection_list = self.coord_input.GetSelectedElement()
                self.event_do(self.user_origin_event,EventOrigin.SELECTIONCOMPLETE_SINGLE)
                return True

        if self.get_selection_mode() == SelectionType.MULTISELECTION:
            self.user_mulitselection_list = self.user_selection.GetSelectedElements(self.coord_input.GetInputViewDocument())
            if len(self.user_mulitselection_list) == 0:
                self.start_selection(SelectionType.MULTISELECTION, self.user_filter, self.user_message)
                return True
            else:
                self.event_do(self.user_origin_event,EventOrigin.SELECTIONCOMPLETE_MULTI)
                return True

        if self.get_selection_mode() == SelectionType.INPUT_POINT:
            input_pnt = self.coord_input.GetInputPoint(mouse_msg, pnt, msg_info)

        if not self.coord_input.IsMouseMove(mouse_msg) and self.get_selection_mode() == SelectionType.INPUT_POINT:
            self.user_referencepoints.append(input_pnt.GetPoint())
            self.event_do(self.user_origin_event,EventOrigin.SELECTIONCOMPLETE_POINT)
            return True

        if self.coord_input.IsMouseMove(mouse_msg):
            return True
        return True

    def start_selection(self,
                        selection_type: SelectionType,
                        filter        : AllplanIFW.SelectionQuery,
                        user_message  : str):
        """ start the selection process defined by a few variables

        Args:
            selection_type: The required selection type, either single, multi or point select.<br>
            filter:         An optional filter to be defined in the selection. Warning: Filter can be overridden in Allplan by user.<br>
            user_message:   Bottom left message to show in Allplan while selecting.
        """
        self.user_filter = filter
        self.user_message = user_message
        self.set_selection_mode(selection_type)

        if selection_type == SelectionType.SINGLE_SELECTION:
            prompt_msg = AllplanIFW.InputStringConvert(user_message)
            self.coord_input.InitFirstElementInput(prompt_msg)

        if filter:
            ele_select_filter = AllplanIFW.ElementSelectFilterSetting(filter,bSnoopAllElements = False)
            self.coord_input.SetElementFilter(ele_select_filter)

        if selection_type == SelectionType.MULTISELECTION:
            AllplanIFW.InputFunctionStarter.StartElementSelect(user_message,ele_select_filter,self.user_selection,markSelectedElements = True)

        if selection_type == SelectionType.INPUT_POINT:
            input_mode = AllplanIFW.CoordinateInputMode(
            identMode       = AllplanIFW.eIdentificationMode.eIDENT_POINT,
            drawPointSymbol = AllplanIFW.eDrawElementIdentPointSymbols.eDRAW_IDENT_ELEMENT_POINT_SYMBOL_YES)
            prompt_msg = AllplanIFW.InputStringConvert(user_message)
            if not self.is_second_input_point:
                self.coord_input.InitFirstPointInput(prompt_msg, input_mode)
                self.is_second_input_point = True
            else:
                self.coord_input.InitNextPointInput(prompt_msg,input_mode)
                self.is_second_input_point = False
        if selection_type == SelectionType.NONE:
            self.coord_input.InitFirstElementInput(AllplanIFW.InputStringConvert("Execute by button click"))

        return

    def event_do(self,
                 event       : Event,
                 event_origin: EventOrigin):
        """ Start or continue an event

        Args:
            event:        Main event from Events Enum<br>
            event_origin: Origin from where the event is fired defined by EventOrigin Enum
        """
        self.set_selection_mode(SelectionType.NONE)
        if event == Event.USER_CONFIRM_EXPORT:
            if(event_origin == EventOrigin.BUTTONCLICK):
                AllplanHelpers.finite_progressbar_create(len(self.created_rebar), "processing", "")
                # write everything to Allplan
                ok, err_msg = AllplanHelpers.write_attributes_to_allplan(self.created_rebar, self.build_ele_list[0].CheckBoxTimestampAttribute.value)
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_ATTRIBUTES_ASSIGNMENT_FAILED) + "\n" + err_msg, AllplanUtil.MB_OK)
                    return None

                # create an IFC file if necessary
                if(self.build_ele_list[0].CheckBoxCreateIFC.value == 1):
                    ifc_path = self.build_ele_list[0].filepathIfc.value
                    ifc_theme = self.build_ele_list[0].IfcExportTheme.value
                    ifc_drawingfiles = self.get_ifc_export_drawing_files()

                    if(not os.path.exists(os.path.dirname(ifc_path))):
                        AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_IFC_PATH_INVALID), AllplanUtil.MB_OK)
                        return None

                    ok = AllplanHelpers.export_ifc_data(ifc_drawingfiles, ifc_path, AllplanBaseElements.IFC_Version.Ifc_4, ifc_theme)
                    if(not ok):
                        AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_IFC_EXPORT_FAILED), AllplanUtil.MB_OK)
                        return None

                self.build_ele_list[0].text_info_user.value = "OK"
                self.ctrl_prop_util.set_enable_function("yesbutton", self.disable_variable_function)
                AllplanHelpers.finite_progressbar_stop()
                return True

        if event == Event.USER_START_EXPORT:
            if(event_origin == EventOrigin.BUTTONCLICK):
                AllplanHelpers.show_message_in_taskbar(AllplanHelpers.get_message(BMWizardInfo.INFO_PREPARING_DATA))

                # get user preferences
                ok, self.attribute_settings = AllplanHelpers.get_user_attribute_settings(self.build_ele_list[0])
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_ATTRIBUTES_UNDEFINED_IN_UI), AllplanUtil.MB_OK)
                    return None

                # select all elements in the drawing
                ok, self.selected_elements = AllplanHelpers.select_drawing_elements()
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_SELECTION_FAILED), AllplanUtil.MB_OK)
                    return None

                # check for assemblies and save the information in a table
                # if there are no assemblies, it will just generate an empty list.
                self.assembly_match_table = AllplanHelpers.get_assembly_information_from_selection(self.selected_elements)

                # filter for rebar elements in selection
                ok, self.selected_elements = AllplanHelpers.filter_drawing_elements_for_rebar(self.selected_elements)
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_INVALID_IFC_TYPE_OR_MISSING), AllplanUtil.MB_OK)
                    return None

                # export the bending machine files to the TMP Allplan folder
                ok = AllplanHelpers.export_bending_machine_files(AllplanSettings.AllplanPaths.GetUsrPath() + "tmp\\bendtemp.bvbs")
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_GENERAL_EXPORT_BVBS_ERROR), AllplanUtil.MB_OK)
                    return None

                # import the bending machine files again
                ok, imported_bvbs_information = AllplanHelpers.import_bending_machine_files(AllplanSettings.AllplanPaths.GetUsrPath() + "tmp\\bendtemp.bvbs")
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_GENERAL_IMPORT_BVBS_ERROR), AllplanUtil.MB_OK)
                    return None

                # write bvbs data to RebarElements with (most) Allplan attributes assigned
                ok, created_rebar = AllplanHelpers.create_rebar_from_bending_machine_files(imported_bvbs_information, self.attribute_settings)
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_GENERAL_PARSING_ERROR) + "\n" + AllplanHelpers.get_exception_message(created_rebar), AllplanUtil.MB_OK)
                    return None

                # angles and lengths do not have user defined attributes and should be created on the fly.
                ok, created_rebar = AllplanHelpers.set_create_segment_angles_lengths_attributes(created_rebar, self.attribute_settings)
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_CREATING_NEW_ATTRIBUTES), AllplanUtil.MB_OK)
                    return None

                # match rebar elements and allplan data, assembly data needed for correct matching of assembly ID's
                ok, self.created_rebar, unassigned_marks = AllplanHelpers.set_corresponding_elements_on_rebarelements(created_rebar, self.selected_elements, self.assembly_match_table)
                if(not ok):
                    missing_marks = " - ".join(unassigned_marks)
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_MATCHING_ALLPLAN_DATA, missing_marks), AllplanUtil.MB_OK)

                # in case of couplers, adjust bar lengths
                ok = AllplanHelpers.adjust_rebar_lengths_for_bars_with_couplers(created_rebar)
                if(not ok):
                    AllplanUtil.ShowMessageBox(AllplanHelpers.get_message(BMWizardInfo.ERR_COUPLER_MATCHING), AllplanUtil.MB_OK)

                # calculate total amount of rebar in case of assemblies
                self.created_rebar = AllplanHelpers.calculate_total_rebar_amounts_for_assemblies(self.created_rebar, self.attribute_settings)

                # display summary tab
                AllplanHelpers.fill_data_summary_tab(self.ctrl_prop_util, self.build_ele_list[0], None)
                self.set_tab_status_summary()
                self.build_ele_list[0].text_info_user.value = "Waiting for user input"
                return True

    def get_ifc_export_drawing_files(self):
        build_ele = self.build_ele_list[0]
        if build_ele.FilesToExport.value == build_ele.IFC_EXPORT_ACTIVE_FILE:
            export_file_numbers = [AllplanBaseElements.DrawingFileService.GetActiveFileNumber()]

        elif build_ele.FilesToExport.value == build_ele.IFC_EXPORT_ALL_FILES:
            export_file_numbers = [file_index for file_index, _ in AllplanBaseElements.DrawingFileService().GetFileState()]

        else:
            export_file_numbers = [int(item.FileName.split("-", 1)[0]) for item in build_ele.FileList.value if item.ExportState]
        return export_file_numbers

    def on_cancel_function(self):
        self.set_tab_status_startup()
        self.palette_service.close_palette()
        AllplanHelpers.finite_progressbar_stop()
        return True

    def on_preview_draw(self):
        return

    def on_mouse_leave(self):
        return

    def set_event(self, event):
        self.user_origin_event = event

    def get_event(self):
        return self.user_fired_event

    def reset_event(self):
        self.user_fired_event = Event.NO_EVENT

    def set_selection_mode(self, type):
        self.user_selection_mode = type

    def get_selection_mode(self):
        return self.user_selection_mode

    def modify_element_property(self, page, name, value):
        """
        Modify property of element

        Args:
            page:   the page of the property
            name:   the name of the property.
            value:  new value for property.
        """
        # IFC Export file list generation if necessary
        if not self.build_ele_list[0].FileList.value:
            if (file_list_tuple := BuildingElementTupleUtil.create_namedtuple_from_definition(self.build_ele_list[0].FileList)) is not None:
                self.build_ele_list[0].FileList.value = [  \
                    file_list_tuple(AllplanElementAdapter.DocumentNameService.GetDocumentNameByFileNumber(index, True, False, "-"), True) \
                    for index, _ in AllplanBaseElements.DrawingFileService().GetFileState()]

        # update palette if necessary
        update_palette = self.palette_service.modify_element_property(page, name, value)

        if update_palette:
            self.palette_service.update_palette(-1, False)

    def execute_load_favorite(self, file_name):
        """ load the favorite data """

        BuildingElementListService.read_from_file(file_name, self.build_ele_list)

        self.palette_service.update_palette(-1, True)

    def reset_param_values(self, _build_ele_list):
        """ reset the parameter values """

        BuildingElementListService.reset_param_values(self.build_ele_list)
        AllplanHelpers.finite_progressbar_stop()
         # workaround "list may not be empty upon visibility change" BUG
        temp_list = self.build_ele_list[0].AnyValueByTypeList.value
        temp_list.append(AnyValueByType.AnyValueByType("Text", " ", ""))
        # update palette
        self.palette_service.update_palette(-1, True)

    def update_after_favorite_read(self):
        """
        Update the data after a favorite read
        """

        self.palette_service.update_palette(-1, True)

    def __del__(self):
        BuildingElementListService.write_to_default_favorite_file(self.build_ele_list)

    def set_active_palette_page_index(self, active_page_index: int):
        self.palette_service.update_palette(-1, False)

    def set_tab_status_startup(self):
        self.build_ele_list[0].is_start_visible.value = 1
        self.build_ele_list[0].is_summary_visible.value = 0

    def set_tab_status_summary(self):
        self.build_ele_list[0].is_start_visible.value = 0
        self.build_ele_list[0].is_summary_visible.value = 1


class ReportHelper():
    reportdata = []

    @staticmethod
    def save(name, value):
        ReportHelper.reportdata.append(ReportElement(name,value))

    @staticmethod
    def reset():
        ReportHelper.reportdata = []

    def get():
        return ReportHelper.reportdata


class ReportElement():

    def __init__(self, name, value):
        self.name = name
        self.value = value


class AllplanHelpers():
    """Contains all helper methods to run the program.
    - most helper methods are self explanatory. methods preceded with __ are internal and should not be used outside of the Allplanhelper construct
    - keeps track of most of the data in the program
    """
    coord_input = None
    doc = None
    string_table = None
    first_run = True # identifier for progress bar if it needs to be created or a step needs to be set.
    progress_bar_finite = None

    @staticmethod
    def calculate_total_rebar_amounts_for_assemblies(rebar_elements, attribute_preferences):
        set_marks = []
        for reb in rebar_elements:
            if reb.is_part_of_assembly:
                set_marks.append(reb.mark.value)
        set_marks = set(set_marks)
        amounts = []
        for mark in set_marks:
            temp_amount = 0
            for reb in rebar_elements:
                if reb.mark.value == mark and reb.is_part_of_assembly:
                    temp_amount = temp_amount + int(reb.amount_assembly.value)
            for index, reb in enumerate(rebar_elements):
                if reb.mark.value == mark and reb.is_part_of_assembly:
                    rebar_elements[index].amount_total = RebarElementAttribute(attribute_preferences["rebaramounttotal"][0].value, str(temp_amount))

        return rebar_elements

    @staticmethod
    def fill_data_summary_tab(ctrl_prop_util: ControlPropertiesUtil, build_ele, data_list):
        build_ele.AnyValueByTypeList.value = []
        value = build_ele.AnyValueByTypeList.value
        if value == []:
            for report_element in ReportHelper.get():
                value.append(AnyValueByType.AnyValueByType("Text", report_element.name, report_element.value))

    @staticmethod
    def finite_progressbar_create(steps, title, description):
        AllplanHelpers.progress_bar_finite = AllplanUtil.ProgressBar(steps,0,False)
        AllplanHelpers.progress_bar_finite.StartProgressbar(steps, title, description, True, True)
        AllplanHelpers.progress_bar_finite.SetAditionalInfo(title)

    @staticmethod
    def finite_progressbar_stop():
        try:
            AllplanHelpers.progress_bar_finite.CloseProgressbar()
        except:
            pass

    @staticmethod
    def finite_progressbar_step():
        AllplanHelpers.progress_bar_finite.Step()

    @staticmethod
    def log(location: str, message, is_error_message: bool):
        if(is_error_message):
            message = AllplanHelpers.get_exception_message(message)
        print(location + " -> " + message)

    @staticmethod
    def static_init(coord_input, string_table: BuildingElementStringTable):
        AllplanHelpers.coord_input = coord_input
        AllplanHelpers.string_table = string_table
        AllplanHelpers.doc = coord_input.GetInputViewDocument()

    @staticmethod
    def show_message_in_taskbar(message: str):
        AllplanHelpers.coord_input.InitFirstElementInput(AllplanIFW.InputStringConvert(message))
        return

    @staticmethod
    def export_ifc_data(export_file_numbers, file_path, ifc_version, ifc_theme):
        AllplanHelpers.show_message_in_taskbar(AllplanHelpers.get_message(BMWizardInfo.INFO_EXPORT_IFC))
        export_import_service = AllplanBaseElements.ExportImportService()
        try:
            export_import_service.ExportIFC(AllplanHelpers.doc,export_file_numbers, ifc_version, file_path, ifc_theme)
            return True
        except Exception as exc:
            AllplanHelpers.log("export_ifc_data", exc, True)
            return False

    @staticmethod
    def export_bending_machine_files(file_path:str) -> bool:
        if(Path(file_path).is_file()):
            Path.unlink(file_path)
        try:
            AllplanBaseElements.DrawingFileService.ExportBendingMachine(AllplanBaseElements.DrawingFileService(), AllplanHelpers.doc, file_path, "project", "plan", "index", False)
            return True
        except Exception as exc:
            AllplanHelpers.log("export_bending_machine_files", exc, True)
            return False

    @staticmethod
    def import_bending_machine_files(file_path:str) -> tuple[bool, List[str]]:
        temporary_list = None
        try:
            file = open(file_path, "r")
            temporary_list = file.readlines()
            file.close()
            ReportHelper.save("BVBS definition entries", str(len(temporary_list)))
            return True, temporary_list
        except Exception as exc:
            print(AllplanHelpers.get_exception_message(exc))
            return False, temporary_list

    @staticmethod
    def select_drawing_elements():
        selection_elementadapterlist = AllplanBaseElements.ElementsSelectService.SelectAllElements(AllplanHelpers.doc)
        if(selection_elementadapterlist ==  None):
            return False, None
        ReportHelper.save("Elements in drawing", str(len(selection_elementadapterlist)))
        return True, selection_elementadapterlist

    @staticmethod
    def get_assembly_information_from_selection(selection_elementadapterlist: AllplanElementAdapter):
        assembly_selection = []
        for element in selection_elementadapterlist:
            if element.GetDisplayName() == "Assembly":
                assembly_selection.append(element)
        # elements have been selected now get the assembly name attribute
        assembly_matching_table = []
        for assembly in assembly_selection:
            uuids = []
            attributes = assembly.GetAttributes(AllplanBaseElements.eAttibuteReadState.ReadAllAndComputable)
            assembly_name = AllplanHelpers.linear_search(attributes, 507)
            if(assembly_name):
                assembly_name = assembly_name[1]
            child_placements = AllplanElementAdapter.BaseElementAdapterChildElementsService.GetChildElements(assembly, False)
            for placement in child_placements:
                uuid = AllplanHelpers.__get_placement_uuid(placement)
                if uuid:
                    uuids.append(uuid)
            assembly_matching_table.append(AssemblyElement(assembly_name, uuids))
        ReportHelper.save("Assemblies in drawing", str(len(assembly_matching_table)))
        return assembly_matching_table

    @staticmethod
    def filter_drawing_elements_for_rebar(selection_elementadapterlist: AllplanElementAdapter):
        rebar_selection = []
        placement_uuids = [AllplanElementAdapter.BarsLinearPlacement_TypeUUID,
                           AllplanElementAdapter.BarsLinearMultiPlacement_TypeUUID,
                           AllplanElementAdapter.BarsAreaPlacement_TypeUUID,
                           AllplanElementAdapter.BarsSpiralPlacement_TypeUUID,
                           AllplanElementAdapter.BarsCircularPlacement_TypeUUID,
                           AllplanElementAdapter.BarsRotationalSolidPlacement_TypeUUID,
                           AllplanElementAdapter.BarsRotationalPlacement_TypeUUID,
                           AllplanElementAdapter.BarsTangentionalPlacement_TypeUUID,
                           AllplanElementAdapter.BarsEndBendingPlacement_TypeUUID]
        # first get the rebar and save it to a smaller list to work with
        for element in selection_elementadapterlist:
            attributes = element.GetAttributes(AllplanBaseElements.eAttibuteReadState.ReadAllAndComputable)
            ifc_class = AllplanHelpers.linear_search(attributes, 684)
            if(not ifc_class == None):
                ifc_class = ifc_class[1]
                if(ifc_class == "IfcReinforcingBar" and element.GetElementAdapterType().GetGuid() in placement_uuids):
                    rebar_selection.append(element)
        if(len(rebar_selection) == 0):
            return False, None
        else:
            ReportHelper.save("Actual placements", str(len(rebar_selection)))
            return True, rebar_selection

    @staticmethod
    def get_exception_message(exc: Exception) -> str:
        if hasattr(exc, 'message'):
            return exc.Message
        else:
            return exc.args[0]

    @staticmethod
    def get_message(message: BMWizardInfo, data = None):
        msg_number = 9000 + message.value
        msg = AllplanHelpers.string_table.get_string(str(msg_number), "String not found")
        if(data):
            if(type(data) is str):
                msg = msg + " " + data
            else:
                msg = msg + " " + '-'.join(str(x.value) for x in data)
        return msg

    @staticmethod
    def linear_search(data, target):
        for tup in data:
            if target in tup:
                return tup
        return None

    @staticmethod
    def create_rebar_from_bending_machine_files(bvbs_data_lines: List[str], attribute_preferences):
        created_rebar = []
        for data_line in bvbs_data_lines:
            rebar = RebarElement()
            try:
                rebar.init_from_bvbs(data_line, attribute_preferences)
                created_rebar.append(rebar)
            except Exception as exc:
                AllplanHelpers.log("create_rebar_from_bvbs", exc , True)
                return False, exc
        # for the user, give some more information in the report
        bf2d_amount = 0
        bf3d_amount = 0
        for reb in created_rebar:
            if reb.shape_type == ShapeType.SHAPE2D:
                bf2d_amount = bf2d_amount +1
            else:
                bf3d_amount = bf3d_amount +1
        ReportHelper.save("2D rebar shapes", str(bf2d_amount))
        ReportHelper.save("3D rebar shapes", str(bf3d_amount))
        return True, created_rebar

    @staticmethod
    def get_user_attribute_settings(palette: BuildingElement):
        # get the attribute definitions from the palette
        attribute_preferences = {}
        attribute_preferences["rebarmark"] = [palette.BVBSRebarMarkAttribute]
        attribute_preferences["rebarlength"] = [palette.BVBSRebarLengthAttribute]
        attribute_preferences["rebardiameter"] = [palette.BVBSRebarDiameterRealAttribute]
        attribute_preferences["rebarbending"] = [palette.BVBSRebarBendingAttribute]
        attribute_preferences["rebarlengthx"] = [palette.BVBSLengthAttributeNamePrefix]
        attribute_preferences["rebaranglex"] = [palette.BVBSAngleAttributeNamePrefix]
        attribute_preferences["rebarbendx"] = [palette.BVBSBendAttributeNamePrefix]
        attribute_preferences["rebarassembly"] = [palette.BVBSAssemblyAttribute]
        attribute_preferences["rebarcouplerstart"] = [palette.BVBSCouplerStartAttribute]
        attribute_preferences["rebarcouplerstartfabricant"] = [palette.BVBSCouplerStartFabricantAttribute]
        attribute_preferences["rebarcouplerstarttype"] = [palette.BVBSCouplerStartTypeAttribute]
        attribute_preferences["rebarcouplerend"] = [palette.BVBSCouplerEndAttribute]
        attribute_preferences["rebarcouplerendfabricant"] = [palette.BVBSCouplerEndFabricantAttribute]
        attribute_preferences["rebarcouplerendtype"] = [palette.BVBSCouplerEndTypeAttribute]
        attribute_preferences["rebaramounttotal"] = [palette.BVBSRebarAmountTotalAttribute]
        attribute_preferences["rebaramountassembly"] = [palette.BVBSRebarAmountAssemblyAttribute]
        attribute_preferences["rounding"] = [palette.roundingcombobox]
        attribute_preferences["arcradius"] = [palette.BVBSRebarRadiusArcAttribute]


        # check if all attributes are defined
        for attr in attribute_preferences:
            if(str(attribute_preferences[attr][0].value) == "0"): #TODO check if we can include the questionmark in the checking of the files
                return False, None
        return True, attribute_preferences

    @staticmethod
    def create_new_attribute_in_allplan(attribute_name, attribute_type: AllplanBaseElements.AttributeService.AttributeType, attribute_dimension) -> int:
        # We create a new user attribute, however, if it already exists ( name check ) then use that attribute instead.
        attribute_number = AllplanBaseElements.AttributeService.GetAttributeID(AllplanHelpers.doc, attribute_name)
        if(attribute_number == -1):
            attr_list_values =  AllplanUtil.VecStringList()
            attr_control_type = AllplanBaseElements.AttributeService.AttributeControlType.Edit
            attribute_number = AllplanBaseElements.AttributeService.AddUserAttribute(
                                                      doc=                      AllplanHelpers.doc,
                                                      attributeType=            attribute_type,
                                                      attributeName=            attribute_name,
                                                      attributeDefaultValue=    "",
                                                      attributeMinValue=        0.0,
                                                      attributeMaxValue=        50000.0,
                                                      attributeDimension=       attribute_dimension,
                                                      attributeCtrlType=        attr_control_type,
                                                      attributeListValues=      attr_list_values)
        return attribute_number

    @staticmethod
    def get_attribute_type_for_attribute_id(id):
        return AllplanBaseElements.AttributeService.GetAttributeType(AllplanHelpers.doc, id)

    @staticmethod
    def __alphabet(index:int):
        alfa = "abcdefghijklmnopqrstuvwxyz"
        if len(alfa) <= index:
            return "_OVERFLOW_"
        else:
            return alfa[index].upper()

    @staticmethod
    def set_create_segment_angles_lengths_attributes(rebar_elements, attribute_preferences):
        """ Create only new length and angle attributes where necessary. Other attributes are defined in the palette by the user.
            This will only process lengths and angles if the allplan_attribute_id of the segment attributes has been set to "undefined".
            Other accepted values are "placeholder" and will result in the attribute creation being skipped as well as a shift in the sequential letter order.
        Args:
            rebar_elements:        a list of Elements of type RebarElements
            attribute_preferences: the attribute preferences defined by the user in the palette.
        """
        prefix_length = attribute_preferences["rebarlengthx"][0].value
        prefix_angle = attribute_preferences["rebaranglex"][0].value
        prefix_bend = attribute_preferences["rebarbendx"][0].value
        # this assumes the sequential order of lengths (A,C,E...) and angles (B,D,F...) in the lists. If this is not the case, wrong attribute names may be generated.
        # attribute
        try:
            for index, ele in enumerate(rebar_elements):
                new_lengths = []
                new_angles = []
                new_bendingpins = []
                lengths = ele.segment_lengths
                angles = ele.segment_angles
                bendingpins = ele.segment_angles_bendingpins

                for len in lengths:
                        atttribute_id = AllplanHelpers.create_new_attribute_in_allplan(prefix_length + AllplanHelpers.__alphabet(len.allplan_attribute_id), AllplanBaseElements.AttributeService.AttributeType.Double,"mm")
                        new_lengths.append(RebarElementAttribute(atttribute_id, len.value))


                for ang in angles:
                        atttribute_id = AllplanHelpers.create_new_attribute_in_allplan(prefix_angle + AllplanHelpers.__alphabet(ang.allplan_attribute_id), AllplanBaseElements.AttributeService.AttributeType.Double,"deg")
                        new_angles.append(RebarElementAttribute(atttribute_id, ang.value))

                for bend in bendingpins:
                        atttribute_id = AllplanHelpers.create_new_attribute_in_allplan(prefix_bend + AllplanHelpers.__alphabet(bend.allplan_attribute_id), AllplanBaseElements.AttributeService.AttributeType.Double,"mm")
                        if bend.value: # this is none for all skipped bends, so that letters continue.
                            new_bendingpins.append(RebarElementAttribute(atttribute_id, bend.value))

                rebar_elements[index].segment_lengths = new_lengths
                rebar_elements[index].segment_angles = new_angles
                rebar_elements[index].segment_angles_bendingpins = new_bendingpins
            return True, rebar_elements
        except:
            return False, None

    @staticmethod
    def __get_rebar_mark_for_placement(element, only_global_position):
        parent_element = AllplanElementAdapter.BaseElementAdapterParentElementService.GetParentElement(element)
        rebarmark = AllplanElementAdapter.ReinforcementPropertiesReader.GetPositionNumber(parent_element)
        position_data = AllplanReinforcement.BarPositionData(element)
        rebarmark_sub = position_data.GetSubPosition()
        if rebarmark_sub and not only_global_position:
            if not str(rebarmark_sub) == "0":
                return str(rebarmark) + "." + str(rebarmark_sub)
        return rebarmark

    @staticmethod
    def __get_placement_type_for_placement(element):
        return element.GetElementAdapterType().GetGuid()

    @staticmethod
    def __get_assembly_id_for_placement(mark, assembly_match_table):
        for assembly_element in assembly_match_table:
            for ass_mark in assembly_element.rebar_uuids:
                if ass_mark == mark:
                    return assembly_element.assembly_name
        return None

    @staticmethod
    def __get_placement_uuid(element):
        return element.GetElementUUID()

    @staticmethod
    def adjust_rebar_lengths_for_bars_with_couplers(rebar_elements):
        success = True
        for rebar_element in rebar_elements:
            success = rebar_element.adjust_first_last_segment_when_coupler()
        return success

    @staticmethod
    def set_corresponding_elements_on_rebarelements(rebar_elements, allplan_selection, assembly_match_table):
        # check place in polygon rebar that only contains one element, this would mean the rebar has not been unlinked
        # BUG: check will fail under the following conditions:
        # two placements, same mark number, one unlinked, one not. There is no possible way to figure out which was was unlinked, which one was not.
        polygonal_placements_not_unlinked_temp_list = []
        polygonal_placements_not_unlinked_global_marks_temp_list = []
        polygonal_placements_not_unlinked_list = []
        for allplan_rebar in allplan_selection:
            allplan_placement = allplan_rebar.GetElementAdapterType().DisplayName
            if allplan_placement == "Place in polygon":
                allplan_mark = str(AllplanHelpers.__get_rebar_mark_for_placement(allplan_rebar, False))
                polygonal_placements_not_unlinked_temp_list.append(allplan_mark)
                polygonal_placements_not_unlinked_global_marks_temp_list.append(allplan_mark.split(".")[0])
        for placement_global_mark in polygonal_placements_not_unlinked_global_marks_temp_list:
            times_placement_found = []
            for placement_mark in polygonal_placements_not_unlinked_temp_list:
                if placement_mark.startswith(placement_global_mark):
                    times_placement_found.append(placement_mark)
            times_placement_found = list(set(times_placement_found))
            if len(times_placement_found) == 1:
                polygonal_placements_not_unlinked_list.append(placement_global_mark + ".1")

        unassigned_allplan_marks = []
        for allplan_rebar in allplan_selection:
            allplan_mark = AllplanHelpers.__get_rebar_mark_for_placement(allplan_rebar, False)
            allplan_uid = AllplanHelpers.__get_placement_uuid(allplan_rebar)
            assembly_id = AllplanHelpers.__get_assembly_id_for_placement(allplan_uid, assembly_match_table)
            match_is_found = False
            # test for unlinked placements, if not unlinked, then make it fail
            if str(allplan_mark) in polygonal_placements_not_unlinked_list:
                allplan_mark = AllplanHelpers.__get_rebar_mark_for_placement(allplan_rebar, True)
            # iteration for all placements dependent on if assembly has been found or not.
            if not assembly_id:
                for index, rebar_element in enumerate(rebar_elements):
                    if str(rebar_element.mark.value) == str(allplan_mark):
                        AllplanHelpers.__set_corresponding_element_on_rebarelement(rebar_elements[index], allplan_rebar)
                        match_is_found = True
                        break
            else:
                for index, rebar_element in enumerate(rebar_elements):
                    if rebar_element.assembly:
                        if str(rebar_element.mark.value) == str(allplan_mark) and str(assembly_id) == str(rebar_element.assembly.value):
                            AllplanHelpers.__set_corresponding_element_on_rebarelement(rebar_elements[index], allplan_rebar)
                            match_is_found = True
                            break

            # no match is found, we should add unassigned elements to a set.
            if not match_is_found:
                unassigned_allplan_marks.append(str(allplan_mark))
        if(len(unassigned_allplan_marks) > 0):
            ReportHelper.save("Unassigned Allplan Elements", str(len(unassigned_allplan_marks)))
            return False, rebar_elements, unassigned_allplan_marks
        return True, rebar_elements, None

    @staticmethod
    def __set_corresponding_element_on_rebarelement(rebar_element, allplan_selected_element):
        rebar_element.allplan_elements.append(allplan_selected_element)
        allplan_placement_type = AllplanHelpers.__get_placement_type_for_placement(allplan_selected_element)
        rebar_element.allplan_placement_type = allplan_placement_type

    @staticmethod
    def write_attributes_to_allplan(rebar_elements, create_timestamp_attribute):
        current_attribute = "Current Attribute: None"
        try:
            for rebar_element in rebar_elements:
                    AllplanHelpers.finite_progressbar_step()
                    attributes = BuildingElementAttributeList()

                    for attribute in rebar_element.get_attributes_as_list():
                        if not attribute:
                            continue
                        current_attribute = "Current Attribute: attribute id: " + str(attribute.allplan_attribute_id) + " & value: " + str(attribute.value)

                        attribute_type = AllplanHelpers.get_attribute_type_for_attribute_id(attribute.allplan_attribute_id)

                        if(attribute_type == AllplanBaseElements.AttributeService.String):
                            attributes.add_attribute_by_unit(int(attribute.allplan_attribute_id), str(attribute.value))
                        elif(attribute_type == AllplanBaseElements.AttributeService.Double):
                            attributes.add_attribute_by_unit(int(attribute.allplan_attribute_id), float(attribute.value))
                        elif(attribute_type == AllplanBaseElements.AttributeService.Integer):
                            attributes.add_attribute_by_unit(int(attribute.allplan_attribute_id), int(attribute.value))


                    if(create_timestamp_attribute):
                        try:
                            attributes.add_attribute(27553, str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
                        except:
                            pass

                    attr_list = attributes.get_attributes_list_as_tuples()
                    element_list = AllplanElementAdapter.BaseElementAdapterList()
                    for allplan_element in rebar_element.allplan_elements:
                        element_list.append(allplan_element)
                    AllplanBaseElements.ElementsAttributeService.ChangeAttributes(attr_list, element_list)

            return True, None
        except:
            return False, current_attribute

    @staticmethod
    def round(value, user_preference):
        round_value = int(user_preference)
        calc_value = int(value)
        return round(calc_value / round_value) * round_value


class RebarElement():
    """An Element of Rebar container which is used throughout the program to contain all information necessary to execution:
    - the values of the attributes
    - the shape information
    - the placement information
    - segment angles and lengths
    - the associated Allplan objects
    """
    def __init__(self) -> None:
        self.shape_type = None
        self.allplan_elements = []
        self.allplan_placement_type = None
        self.geometry_type = None
        self.mark = None
        self.is_circular_reinforcement = False
        self.is_part_of_assembly = False
        self.total_length = None
        self.diameter = None
        self.bend_angle = None
        self.assembly = None
        self.coupler_start = None
        self.coupler_end = None
        self.coupler_start_fabricant = None
        self.coupler_start_type = None
        self.coupler_end_fabricant = None
        self.coupler_end_type = None
        self.segment_lengths = []
        self.segment_angles = []
        self.amount_total = None
        self.amount_assembly = None
        self.radius = None
        self.segment_angles_bendingpins = []


    def adjust_first_last_segment_when_coupler(self):
        # Ensure we have Allplan elements and at least one coupler flag enabled
        if not self.allplan_elements or not (self.coupler_start or self.coupler_end):
            return False

        # Get the fixtures from the first Allplan element
        fixtures = AllplanElementAdapter.BaseElementAdapterChildElementsService.GetChildElements(self.allplan_elements[0], True)
        if not fixtures:
            return False

        # Retrieve the fixture with display name 'Symbol fixture'
        fixture = next((obj for obj in fixtures if obj.GetDisplayName() == 'Symbol fixture'), None)
        if not fixture:
            return False

        # Get the attributes of the fixture
        attributes = AllplanBaseElements.ElementsAttributeService.GetAttributes(
            fixture, AllplanBaseElements.eAttibuteReadState.ReadAllAndComputable
        )
        if not attributes:
            return False

        # Retrieve the fixture length and convert it to a float
        fixture_length = AllplanHelpers.linear_search(attributes, 1238)
        if not fixture_length:
            return False
        fixture_length = round(float(fixture_length[1]))

        # Ensure segment_lengths is not empty and convert each RebarElementAttribute.value to a float if needed
        if not self.segment_lengths:
            return False
        for segment in self.segment_lengths:
            if isinstance(segment.value, str):
                try:
                    segment.value = float(segment.value)
                except ValueError:
                    return False  # Conversion failed; handle as needed

        # Adjust segments based on the coupler flags
        if len(self.segment_lengths) == 1:
            # For a single segment, add fixture_length for each enabled coupler flag
            if self.coupler_start.value == "True":
                self.segment_lengths[0].value += fixture_length
            if self.coupler_end.value == "True":
                self.segment_lengths[0].value += fixture_length
        else:
            # For multiple segments, update the first and/or last segment as needed
            if self.coupler_start.value == "True":
                self.segment_lengths[0].value += fixture_length
            if self.coupler_end.value == "True":
                self.segment_lengths[-1].value += fixture_length

        return True


    def __init_bvbs_geometry(self, bvbs_geometry, geometry_type, attribute_preferences):
        if geometry_type == ShapeType.SHAPE2D:
            upper_limit = 400
            rebar_diameter = 10
            length_counter = 0
            angle_counter = 1
            is_first_angle_and_length_after_arc = False # Allplan BUG l=0 values skip
            for index, geo in enumerate(bvbs_geometry):
                    if(geo.startswith("l")):
                        geo_value = geo.replace("l","",1)
                        if geo_value == "0":
                            continue
                        self.segment_lengths.append(RebarElementAttribute(length_counter, geo_value))
                        length_counter+=2

                    if(geo.startswith("w")):
                        if is_first_angle_and_length_after_arc:
                            #self.segment_angles.append(RebarElementAttribute(angle_counter, "0")) # add a zero angle because it is missing
                            is_first_angle_and_length_after_arc = False
                        geo_value = geo.replace("w","",1)
                        self.segment_angles.append(RebarElementAttribute(angle_counter, geo_value))
                        angle_counter+=2

                    if(geo.startswith("r")):
                        geo_value = geo.replace("r","",1)
                        if float(geo_value) > upper_limit:
                            try: # case arc created by user and this is the r value describing the arc (defined by upper limit)
                                arc_radius = float(bvbs_geometry[index].replace("r","",1)) # radius starts with r
                                self.radius = RebarElementAttribute(attribute_preferences["arcradius"][0].value, str(arc_radius))
                                arc_angle = float(bvbs_geometry[index+1].replace("w","",1)) # BVBS definition r should be followed with w
                                arc_length = str(self.__calculate_arc_length(arc_radius, arc_angle))
                                self.segment_lengths.append(RebarElementAttribute(length_counter, arc_length))
                                length_counter+=2
                                is_first_angle_and_length_after_arc = True
                            except:
                                Exception("[Exception] Circular shape does not contain required parameters in BVBS")
                        else: # case fake bending pin created by user and this is the r value describing the bending pin
                            bp_radius = float(geo.replace("r","",1)) # bending pin radius
                            bending_pin = ( bp_radius *2 ) / rebar_diameter
                            self.segment_angles_bendingpins.append(RebarElementAttribute(angle_counter, bending_pin))

        else: # ShapeType.SHAPE3D
            vectors = []
            temp_x = None
            temp_y = None
            temp_z = None
            for geo in bvbs_geometry:
                if(geo.startswith("x")):
                    temp_x = geo.replace("x","")
                if(geo.startswith("y")):
                    temp_y = geo.replace("y","")
                if(geo.startswith("z")):
                    temp_z = geo.replace("z","")
                if temp_z and temp_y and temp_x:
                    vectors.append(Vector(Point(temp_x, temp_y, temp_z)))
                    temp_x = None
                    temp_y = None
                    temp_z = None

            # starting point is 0,0,0
            points = []
            pt1 = Point(0,0,0)
            points.append(Point(0,0,0))
            for vector in vectors:
                pt1 = pt1.move(vector)
                points.append(Point(pt1.x, pt1.y, pt1.z))

            # calculate lengths ( segments = # points-1)
            temp_length_list = []
            for index in range(len(points) - 1):
                pt1 = points[index]
                pt2 = points[index + 1]
                distance = pt1.distance(pt2)
                self.segment_lengths.append(RebarElementAttribute("undefined", str(distance)))

            # calculate angles (angles = # points-2)
            for index in range(len(points) - 2):
                pt1 = points[index]
                pt2 = points[index + 1]
                pt3 = points[index + 2]
                v1 = Vector(pt1, pt2)
                v2 = Vector(pt2, pt3)
                angle = v1.angle_with(v2)
                self.segment_angles.append(RebarElementAttribute("undefined", str(angle)))

        # check if the last value is zero. In that case it should be removed.
        if len(self.segment_angles) > 0:
            if str(self.segment_angles[len(self.segment_angles)-1].value) == "0":
                self.segment_angles.pop()

    def __split_string_at_capitals(self, input_string):
        # Use regex to find the pattern '@' followed by an uppercase letter
        pattern = r'(@[A-Z])'
        parts = re.split(pattern, input_string)
        return parts

    def __calculate_arc_length(self, radius, angle):
        pi = np.pi
        arc_length = (2 * pi * radius) * (angle / 360)
        return arc_length

    def init_from_bvbs(self, data_line : str, attribute_preferences):

        if("BF2D@" in data_line):
            self.shape_type = ShapeType.SHAPE2D
        elif("BF3D@" in data_line):
            self.shape_type = ShapeType.SHAPE3D
        else:
            raise Exception(" [Exception] Unsupported shape: " + data_line)

        bvbs_parts = self.__split_string_at_capitals(data_line)
        bvbs_coupler = None
        bvbs_header = None
        bvbs_geometry = None
        bvbs_assembly = None

        try:
            bvbs_coupler_index = bvbs_parts.index("@M")
            bvbs_coupler = bvbs_parts[bvbs_coupler_index+1]
        except ValueError:
            pass # no couplers

        try:
            bvbs_assembly_index = bvbs_parts.index("@P")
            bvbs_assembly = bvbs_parts[bvbs_assembly_index+1]
            self.is_part_of_assembly = True
        except ValueError:
            pass # no assembly

        try:
            bvbs_geometry_index = bvbs_parts.index("@G")
            bvbs_header_index = bvbs_parts.index("@H")
            bvbs_header = bvbs_parts[bvbs_header_index+1]
            bvbs_geometry = bvbs_parts[bvbs_geometry_index+1]
        except ValueError:
            raise Exception(" [Exception] Syntax error in data: " + data_line + "\nHeader information missing!") # this is not fine

        # now we have the data ordered in an easy line per type

        ### HEADER ###
        bvbs_header = bvbs_header.split("@")
        for bvbs_element in bvbs_header:
            if(bvbs_element.startswith("p")):
                bvbs_value = bvbs_element.replace("p","",1)
                self.mark = RebarElementAttribute(attribute_preferences["rebarmark"][0].value, bvbs_value)
            if(bvbs_element.startswith("l")):
                bvbs_value = bvbs_element.replace("l","",1)
                bvbs_value = str(AllplanHelpers.round(bvbs_value, attribute_preferences["rounding"][0].value))
                self.total_length = RebarElementAttribute(attribute_preferences["rebarlength"][0].value, bvbs_value)
            if(bvbs_element.startswith("d")):
                bvbs_value = bvbs_element.replace("d","",1)
                self.diameter = RebarElementAttribute(attribute_preferences["rebardiameter"][0].value, bvbs_value)
            if(bvbs_element.startswith("s")):
                bvbs_value = bvbs_element.replace("s","",1)
                self.bend_angle = RebarElementAttribute(attribute_preferences["rebarbending"][0].value, bvbs_value)
            if(bvbs_element.startswith("n")):
                bvbs_value = bvbs_element.replace("n","",1)
                if self.is_part_of_assembly:
                    self.amount_assembly = RebarElementAttribute(attribute_preferences["rebaramountassembly"][0].value, bvbs_value)
                else:
                    self.amount_total = RebarElementAttribute(attribute_preferences["rebaramounttotal"][0].value, bvbs_value)

        ### ASSEMBLY ###
        if bvbs_assembly:
            bvbs_assembly = bvbs_assembly.split("@")
            for bvbs_element in bvbs_assembly:
                if(bvbs_element.startswith("t")):
                    bvbs_value = bvbs_element.replace("t","",1)
                    self.assembly = RebarElementAttribute(attribute_preferences["rebarassembly"][0].value, bvbs_value)

        ### COUPLERS ###
        if bvbs_coupler:
            bvbs_coupler = bvbs_coupler.split("@")
            for bvbs_element in bvbs_coupler:
                if(bvbs_element.startswith("c")):
                    bvbs_value = bvbs_element.replace("c","",1)
                    if(bvbs_value == "1"):
                        self.coupler_start = RebarElementAttribute(attribute_preferences["rebarcouplerstart"][0].value, "True")
                    else:
                        self.coupler_start = RebarElementAttribute(attribute_preferences["rebarcouplerstart"][0].value, "False")
                if(bvbs_element.startswith("p")):
                    bvbs_value = bvbs_element.replace("p","",1)
                    if(bvbs_value == "1"):
                        self.coupler_end = RebarElementAttribute(attribute_preferences["rebarcouplerend"][0].value, "True")
                    else:
                        self.coupler_end = RebarElementAttribute(attribute_preferences["rebarcouplerend"][0].value, "False")
                if(bvbs_element.startswith("a")):
                    bvbs_value = bvbs_element.replace("a","",1)
                    if(not bvbs_value.isdigit()):
                        self.coupler_start_fabricant = RebarElementAttribute(attribute_preferences["rebarcouplerstartfabricant"][0].value, bvbs_value)
                if(bvbs_element.startswith("b")):
                    bvbs_value = bvbs_element.replace("b","",1)
                    self.coupler_start_type = RebarElementAttribute(attribute_preferences["rebarcouplerstarttype"][0].value, bvbs_value)
                if(bvbs_element.startswith("n")):
                    bvbs_value = bvbs_element.replace("n","",1)
                    if(not bvbs_value.isdigit()):
                        self.coupler_end_fabricant = RebarElementAttribute(attribute_preferences["rebarcouplerendfabricant"][0].value, bvbs_value)
                if(bvbs_element.startswith("o")):
                    bvbs_value = bvbs_element.replace("o","",1)
                    self.coupler_end_type = RebarElementAttribute(attribute_preferences["rebarcouplerendtype"][0].value, bvbs_value)

        ### GEOMETRY ###
        bvbs_geometry = bvbs_geometry.split("@")
        self.__init_bvbs_geometry(bvbs_geometry, self.shape_type, attribute_preferences)

    def get_attributes_as_list(self):
        attributes_list = []
        attributes_list.append(self.mark)
        attributes_list.append(self.total_length)
        attributes_list.append(self.diameter)
        attributes_list.append(self.bend_angle)
        attributes_list.append(self.assembly)
        attributes_list.append(self.coupler_start)
        attributes_list.append(self.coupler_end)
        attributes_list.append(self.coupler_start_fabricant)
        attributes_list.append(self.coupler_end_type)
        attributes_list.append(self.coupler_start_type)
        attributes_list.append(self.coupler_end_fabricant)
        attributes_list.append(self.amount_total)
        attributes_list.append(self.amount_assembly)
        attributes_list.append(self.radius)
        for segment_length in self.segment_lengths:
            attributes_list.append(segment_length)
        for segment_angle in self.segment_angles:
            attributes_list.append(segment_angle)
        for segment_angle_bendingpin in self.segment_angles_bendingpins:
            attributes_list.append(segment_angle_bendingpin)
        return attributes_list


class Vector():
    """A 3D vector, only used to calculate the vectors of 3D rebar
    contains functions to calculate the angle between two vectors.
    """
    def __init__(self, point_a, point_b = None):
        if point_b:
            self.x = point_b.x - point_a.x
            self.y = point_b.y - point_a.y
            self.z = point_b.z - point_a.z
        else:
            self.x = point_a.x
            self.y = point_a.y
            self.z = point_a.z
        self.array = np.array([self.x, self.y, self.z])

    def __dot_product(self, other_vector):
        return np.dot(self.array, other_vector.array)

    def magnitude(self):
        return np.linalg.norm(self.array)

    def angle_with(self, other_vector):
        dot_prod = self.__dot_product(other_vector)
        mag_v1 = self.magnitude()
        mag_v2 = other_vector.magnitude()
        cos_theta = dot_prod / (mag_v1 * mag_v2)
        angle_rad = np.arccos(cos_theta)
        angle_deg = np.degrees(angle_rad)
        return round(angle_deg)


class Point():
    """A 3D point
    - point can be moved by means of a vector. Will save movement into the point (self)
    - point distance with another point can be calculated with the distance method.
    """
    def __init__(self, x,y,z):
        self.x = int(x)
        self.y = int(y)
        self.z = int(z)

    def move(self, vector):
        self.x = self.x + vector.x
        self.y = self.y + vector.y
        self.z = self.z + vector.z
        return self

    def distance(self, other_point):
        x1 = (self.x)
        y1 = (self.y)
        z1 = (self.z)
        x2 = (other_point.x)
        y2 = (other_point.y)
        z2 = (other_point.z)
        distance = math.sqrt((x2 - x1)**2 + (y2 - y1)**2 + (z2 - z1)**2)
        return round(distance)


class RebarElementAttribute():
    """An container that contains one single attribute of a Rebar Element in Allplan
    The attribute container is populated by a value through BVBS.
    - the allplan attribute ID is the ID of the associated attribute in Allplan
    - the value is the value to be written to Allplan.
    """
    def __init__(self, allplan_attribute_id, allplan_value_to_write):
        self.allplan_attribute_id = allplan_attribute_id
        self.value = allplan_value_to_write


class AssemblyElement():
    """A container to save the Allplan assembly association information
    - the name of the assembly
    - the unique ID's of the Allplan ElementAdapter objects associated with this assembly
    """
    def __init__(self, assembly_name, rebar_uuids):
        self.assembly_name = assembly_name
        self.rebar_uuids = rebar_uuids
