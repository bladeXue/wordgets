import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW, LEFT, CENTER, RIGHT, HIDDEN, VISIBLE
from typing import Union
import platform
import distro
import os
import pyautogui
import webbrowser
import json
import sqlite3
from functools import partial
import re
from datetime import datetime
import copy
import random

# Lib for validating file type
from openpyxl import load_workbook
from bs4 import BeautifulSoup

#Memo 
from supermemo2 import SMTwo

# Globals
g_strRunPath = os.path.dirname(__file__)
g_strLangPath = os.path.join(g_strRunPath, "resources/languages.json")
g_strDBPath = os.path.join(g_strRunPath, "resources/wordgets_stat.db")
g_strConfigPath = os.path.join(g_strRunPath, "resources/configuration.json")
g_strWordlistsPath = os.path.join(g_strRunPath, "resources/wordlists.json")
g_strCardsPath = os.path.join(g_strRunPath, "resources/cards.json")
g_bIsMobile = True
g_EmptyHtml = '<!DOCTYPE html> <html> <head> <title> </title> </head> <body> </body> </html>' #Cannot be "". Otherwise, it will not work

# Whether the form is opened
g_bSubFormReleased = True
g_bSubSubFormReleased = True

g_dictWordLists = {}
g_dictCards     = {}

# Dialog 
g_strMsgBox_OpenFileDialogTitle = ""
g_strMsgBox_ErrorDialogTitle = ""
g_strMsgBox_ErrorDialogMsg_InvalidFile = ""
g_strMsgBox_InfoDialogTitle = ""
g_strMsgBox_InfoDialogMsg_EmptyCardName = ""
g_strMsgBox_InfoDialogMsg_RepetitiveCardName = ""
g_strMsgBox_InfoDialogMsg_EmptyWordListName = ""
g_strMsgBox_InfoDialogMsg_RepetitiveWordListName = ""
g_strMsgBox_ErrorDialogMsg_CannotCreateCard = ""
g_strMsgBox_ErrorDialogMsg_CannotCreateWordlist = ""
g_strMsgBox_ErrorDialogMsg_CardHasAssociatedWordlist = ""
g_strMsgBox_ErrorDialogMsg_Card1IsNull = ""
g_strMsgBox_QuestionDialogTitle_Warning = ""
g_strMsgBox_QuestionDialogMsg_WarningOnDeletingWordList = ""
g_strMsgBox_QuestionDialogMsg_NotFoundWordList = ""
g_strMsgBox_ErrorDialogMsg_NoWordList = ""
g_strMsgBox_ErrorDialogMsg_Card1SameAsCard2 = ""
g_strMsgBox_InfoDialogMsg_Completed = ""
g_strMsgBox_QuestionDialogMsg_DeleteThisWord= ""
g_strMsgBox_InfoDialogMsg_ContinueToStudyNewWords = ""
g_strMsgBox_InfoDialogMsg_MissionComplete = ""
g_strMsgBox_ErrorDialogMsg_InvalidFiles = ""

# Keywords for text display
g_strFront = ""
g_strBack = ""
g_strOnesided = ""
g_strFrontend =""
g_strBackend =""
g_strFilePath = ""
g_strCard1 = ""
g_strCard2 = ""

#Window widgets
g_strLblNewWordsPerGroup = ""
g_strLblOldWordsPerGroup = ""
g_strLblReviewPolicy = ""
g_strCboReviewPolicy_item_LearnFirst = ""
g_strCboReviewPolicy_item_Random = ""
g_strCboReviewPolicy_item_ReviewFirst = ""

# Config
g_strDefaultLang = "English"
g_strOpenedWordListName = ""
g_bOnlyReview = False

# Temporary variables
g_strTmpNameTxt = ""
g_iTmpPrevWordNo = -1
g_iTmpPrevCardType = -1
g_bSwitchContinueNewWords = True
g_iTmpCntStudiedNewWords = 0
g_iTmpCntStudiedOldWords = 0 # It can be over than real total of old words.

# Functions
def GetOS() -> str:
    strOS = platform.system() # If Windows, return Windows
    if strOS == 'Linux':
        strLinuxDist = distro.linux_distribution() #[!] It is said that possibly there could be another outputs on different Android third-party distributions.
        if strLinuxDist == 'Android':
            strOS = strLinuxDist
    elif strOS == 'Darwin':
        if platform.mac_ver()[0].startswith('iPhone'):
            strOS = 'iOS'
        else:
            strOS = 'MacOS'
    return strOS

def GetRenderedHTML(strHTMLTemplateFilePath: str, strPyFilePath: str, dicDynVariables: str) -> str:
    folder_path, file_name = os.path.split(strHTMLTemplateFilePath)
    dicDynVariables['folder_path'] = folder_path
    dicDynVariables['file_name'] = file_name
    with open(strPyFilePath, "r",encoding="utf-8") as ofs:
        exec(ofs.read(), globals(),dicDynVariables)
    strRenderedHTML = dicDynVariables.get("output")
    return strRenderedHTML

def ConvertNumToExcelColTitle(n: int) -> str: #Use for convert a number to the Excel column serial name. Pay attention: from 1
    column_title = ""
    while n > 0:
        remainder = (n - 1) % 26
        column_title = chr(65 + remainder) + column_title
        n = (n - 1) // 26
    return column_title

def FilterForExcelNoneValue(strInput: Union[str, None]) -> str:
    if strInput == None:
        return ""
    return strInput

class wordgets(toga.App):
    #Hide menu bar
    def _create_impl(self):
        factory_app = self.factory.App
        factory_app.create_menus = lambda _: None
        return factory_app(interface=self)

    def startup(self):
        # Main window layout
        self.main_window    = toga.MainWindow(title=self.formal_name)
        #wordgets.on_exit    = self.cb_on_exit             A bug.
        boxMain_wndMain     = toga.Box(style = Pack(direction = COLUMN))
        boxRow1_wndMain     = toga.Box(style = Pack(direction = ROW, alignment = CENTER))
        self.btnOpenWndOpt_wndMain = toga.Button("🔧", on_press = self.cbOpenWndOpt)
        self.btnOpenWndMgmt_wndMain = toga.Button("📚", on_press = self.cbOpenWndMgmt)
        self.cboCurWordlist_wndMain = toga.Selection(on_select = self.cbChangeCurWordList)
        self.chkIsReviewOnly_wndMain = toga.Switch("",on_change=self.cbChangeChkIsReviewOnly)
        boxRow1_wndMain.add(
            self.btnOpenWndOpt_wndMain,
            self.btnOpenWndMgmt_wndMain,
            self.cboCurWordlist_wndMain,
            self.chkIsReviewOnly_wndMain
        )

        boxRow2_wndMain                         = toga.Box(style = Pack(direction = ROW, alignment=CENTER))
        boxRow2Left_wndMain                     = toga.Box(style = Pack(flex = 5, direction = ROW, alignment = CENTER, padding_top = 5, padding_bottom = 5))
        self.lblLearningProgressItem_wndMain    = toga.Label("")
        self.lblLearningProgressValue_wndMain   = toga.Label("", style = Pack(flex = 1))
        boxRow2Left_wndMain.add(
            self.lblLearningProgressItem_wndMain,
            self.lblLearningProgressValue_wndMain
        )

        self.boxRow2Right_wndMain       = toga.Box(style = Pack(flex = 1, direction = ROW, alignment = CENTER))
        boxRow2_wndMain.add(
            boxRow2Left_wndMain,
            self.boxRow2Right_wndMain
        )
        
        boxRow3_wndMain     = toga.Box(style = Pack(direction = ROW, flex = 1, padding_top = 5, alignment = CENTER))
        self.wbCard_wndMain = toga.WebView(style = Pack(direction = COLUMN, flex = 1))
        boxRow3_wndMain.add(
            self.wbCard_wndMain
        )

        self.boxRow4_wndMain= toga.Box(style=Pack(direction = ROW, padding_top=5, alignment = CENTER))
        
        boxMain_wndMain.add(
            boxRow1_wndMain,
            boxRow2_wndMain,
            boxRow3_wndMain,
            self.boxRow4_wndMain
        )

        self.main_window.content = boxMain_wndMain

        # Option window layout
        self.wndOpt         = toga.Window("", on_close = self.cbCloseWndOpt)
        self.windows.add(self.wndOpt)
        sbMain_wndOpt       = toga.ScrollContainer(horizontal= False)
        
        boxChangeLang_wndOpt= toga.Box(style = Pack(direction = ROW, alignment = CENTER))
        self.lblLang_wndOpt = toga.Label("")
        lblLangEmoji_wndOpt = toga.Label("🌐")
        self.cboChangeLang_wndOpt = toga.Selection(on_select=self.cbChangeLanguageForCboOnSelect)
        boxChangeLang_wndOpt.add(
            self.lblLang_wndOpt,
            lblLangEmoji_wndOpt,
            self.cboChangeLang_wndOpt
        )
        
        boxVisitOfficialWebsite_wndOpt = toga.Box(style = Pack(direction = ROW, padding_top = 5))
        self.btnVisitOfficialWebsite_wndOpt = toga.Button("", on_press = lambda _widget: webbrowser.open('https://github.com/leaffeather/wordgets'))
        boxVisitOfficialWebsite_wndOpt.add(self.btnVisitOfficialWebsite_wndOpt)
        
        sbMain_wndOpt.content = toga.Box(
            children = [
                boxChangeLang_wndOpt,
                boxVisitOfficialWebsite_wndOpt
            ], 
            style = Pack(direction = COLUMN)
        )
        self.wndOpt.content   = sbMain_wndOpt
        
        # Word list manager
        self.wndMgmt          = toga.Window("", on_close=self.cbCloseWndMgmt)
        self.windows.add(self.wndMgmt)
        
        boxMain_wndMgmt       = toga.Box(style = Pack(direction = COLUMN))
        boxCardGroup_wndMgmt       = toga.Box(style = Pack(direction = ROW, flex = 1))
        boxCardCol1_wndMgmt   = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        self.lblCard_wndMgmt  = toga.Label("")
        btnAddCard_wndMgmt    = toga.Button("➕",style = Pack(flex = 1), on_press = self.cbOpenWndAddCard)
        boxCardCol1_wndMgmt.add(
            self.lblCard_wndMgmt,
            btnAddCard_wndMgmt
        )
        self.sbCardCol2_wndMgmt   = toga.ScrollContainer(style = Pack(direction = COLUMN, flex = 9)) # A bug. Waiting for toga release   
        self.boxCards_wndMgmt = toga.Box(style = Pack(direction = COLUMN))
        self.sbCardCol2_wndMgmt.content = self.boxCards_wndMgmt
        boxCardGroup_wndMgmt.add(
            boxCardCol1_wndMgmt,
            self.sbCardCol2_wndMgmt
        )

        boxWordListGroup_wndMgmt       = toga.Box(style = Pack(direction = ROW, flex = 1, padding_top = 5)) 
        boxWordListCol1_wndMgmt   = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        self.lblWordList_wndMgmt  = toga.Label("")
        btnAddWordList_wndMgmt    = toga.Button("➕",style = Pack(flex = 1), on_press=self.cbOpenWndAddWordList)
        boxWordListCol1_wndMgmt.add(
            self.lblWordList_wndMgmt,
            btnAddWordList_wndMgmt
        )
        self.sbWordListCol2_wndMgmt   = toga.ScrollContainer(style = Pack(direction = COLUMN, flex = 9)) # A bug. Waiting for toga release
        self.boxWordLists_wndMgmt = toga.Box(style = Pack(direction = COLUMN))
        self.sbWordListCol2_wndMgmt.content = self.boxWordLists_wndMgmt
        boxWordListGroup_wndMgmt.add(
            boxWordListCol1_wndMgmt,
            self.sbWordListCol2_wndMgmt
        )

        self.btnApply_wndMgmt = toga.Button("", on_press = self.cbApplyChangesToDB)

        boxMain_wndMgmt.add(
            boxCardGroup_wndMgmt,
            boxWordListGroup_wndMgmt,
            self.btnApply_wndMgmt
        )
        
        self.wndMgmt.content  = boxMain_wndMgmt
        
        # Add card window
        self.wndAddCard = toga.Window("", on_close = self.cbCloseWndAddCard)
        self.windows.add(self.wndAddCard)
        sbMain_wndAddCard = toga.ScrollContainer(style = Pack(direction = COLUMN, flex = 1), horizontal = False)

        boxRow0_wndAddCard = toga.Box(style = Pack(direction = COLUMN, flex = 1, padding_top = 5)) #Appended. 
        self.lblSpecifyNewCardName_wndAddCard = toga.Label("")
        self.txtSpecifyNewCardName_wndAddCard = toga.TextInput("", on_change = self.cbValidateNameTxt)
        
        boxRow0_wndAddCard.add(
            self.lblSpecifyNewCardName_wndAddCard,
            self.txtSpecifyNewCardName_wndAddCard
        )

        boxRow1_wndAddCard = toga.Box(style = Pack(direction = COLUMN, flex = 1, padding_top = 5))
        boxRow1Row1_wndAddCard = toga.Box(style = Pack(direction = ROW))
        self.lblFrontFrontend_WndAddCard = toga.Label("")
        boxRow1Row1_wndAddCard.add(
            self.lblFrontFrontend_WndAddCard
        )

        boxRow1Row2_wndAddCard = toga.Box(style = Pack(direction = ROW, padding_top = 5))
        btnAddFrontFrontend_wndAddCard = toga.Button("📂", on_press = self.cbFrontFrontend)
        self.txtFrontFrontendFilePath_wndAddCard = toga.TextInput(style = Pack(flex = 1), readonly = True)
        boxRow1Row2_wndAddCard.add(
            btnAddFrontFrontend_wndAddCard,
            self.txtFrontFrontendFilePath_wndAddCard
        )

        boxRow1Row3_wndAddCard = toga.Box(style = Pack(direction = ROW, padding_top = 5))
        self.lblFrontBackend_WndAddCard = toga.Label("")
        boxRow1Row3_wndAddCard.add(
            self.lblFrontBackend_WndAddCard
        )

        boxRow1Row4_wndAddCard = toga.Box(style = Pack(direction = ROW, padding_top = 5))
        btnAddFrontBackend_wndAddCard = toga.Button("📂", on_press = self.cbFrontBackend)
        self.txtFrontBackendFilePath_wndAddCard = toga.TextInput(style = Pack(flex = 1), readonly = True)
        boxRow1Row4_wndAddCard.add(
            btnAddFrontBackend_wndAddCard,
            self.txtFrontBackendFilePath_wndAddCard
        )

        boxRow1_wndAddCard.add(
            boxRow1Row1_wndAddCard,
            boxRow1Row2_wndAddCard,
            boxRow1Row3_wndAddCard,
            boxRow1Row4_wndAddCard,
        )
        
        boxRow2_wndAddCard = toga.Box(style = Pack(direction = COLUMN, flex = 1, padding_top = 5))
        self.chkEnableBack_wndAddCard = toga.Switch("", value = True, on_change = self.cbEnableBack)
        # There is a bug that if initial value is False, the boxRow3 should be hidden but will still be visible on the condition of visibility = HIDDEN  
        boxRow2_wndAddCard.add(
            self.chkEnableBack_wndAddCard
        )
        
        self.boxRow3_wndAddCard = toga.Box(style = Pack(direction = COLUMN, padding_top = 5, flex = 1))
        boxRow3Row1_wndAddCard = toga.Box(style = Pack(direction = ROW))
        self.lblBackFrontend_WndAddCard = toga.Label("")
        boxRow3Row1_wndAddCard.add(
            self.lblBackFrontend_WndAddCard
        )

        boxRow3Row2_wndAddCard = toga.Box(style = Pack(direction = ROW, padding_top = 5))
        btnAddBackFrontend_wndAddCard = toga.Button("📂", on_press = self.cbBackFrontend)
        self.txtBackFrontendFilePath_wndAddCard = toga.TextInput(style = Pack(flex = 1), readonly = True)
        boxRow3Row2_wndAddCard.add(
            btnAddBackFrontend_wndAddCard,
            self.txtBackFrontendFilePath_wndAddCard
        )

        boxRow3Row3_wndAddCard = toga.Box(style = Pack(direction = ROW, padding_top = 5))
        self.lblBackBackend_WndAddCard = toga.Label("")
        boxRow3Row3_wndAddCard.add(
            self.lblBackBackend_WndAddCard
        )

        boxRow3Row4_wndAddCard = toga.Box(style = Pack(direction = ROW, padding_top = 5))
        btnAddBackBackend_wndAddCard = toga.Button("📂", on_press = self.cbBackBackend)
        self.txtBackBackendFilePath_wndAddCard = toga.TextInput(style = Pack(flex = 1), readonly = True)
        boxRow3Row4_wndAddCard.add(
            btnAddBackBackend_wndAddCard,
            self.txtBackBackendFilePath_wndAddCard
        )
        self.boxRow3_wndAddCard.add(
            boxRow3Row1_wndAddCard,
            boxRow3Row2_wndAddCard,
            boxRow3Row3_wndAddCard,
            boxRow3Row4_wndAddCard,
        )

        sbMain_wndAddCard.content = toga.Box(
            children = [
                boxRow0_wndAddCard,
                boxRow1_wndAddCard,
                boxRow2_wndAddCard,
                self.boxRow3_wndAddCard
            ], 
            style = Pack(direction = COLUMN)
        )
        self.btnSaveNewCard_wndAddCard = toga.Button("", style = Pack(flex = 1), on_press = self.cbSaveNewCard)

        self.wndAddCard.content = toga.Box(children = [sbMain_wndAddCard, self.btnSaveNewCard_wndAddCard], style = Pack(direction = COLUMN))
        
        # Add word list window
        self.wndAddWordList  = toga.Window("", on_close = self.cbCloseWndAddWordlist)
        self.windows.add(self.wndAddWordList)
        sbMain_wndAddWordList = toga.ScrollContainer(style = Pack(direction = COLUMN, flex = 1), horizontal = False)
        
        boxRow1_wndAddWordList = toga.Box(style = Pack(direction = COLUMN, flex = 1, padding_top = 5)) #Appended. 
        self.lblSpecifyNewWordListName_wndAddWordList = toga.Label("")
        self.txtSpecifyNewWordListName_wndAddWordList = toga.TextInput("", on_change = self.cbValidateNameTxt)
        boxRow1_wndAddWordList.add(
            self.lblSpecifyNewWordListName_wndAddWordList,
            self.txtSpecifyNewWordListName_wndAddWordList
        )

        boxRow2_wndAddWordList = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        self.lblExcelFile_WndAddWordList = toga.Label("")
        boxRow2_wndAddWordList.add(
            self.lblExcelFile_WndAddWordList
        )

        boxRow3_wndAddWordList = toga.Box(style = Pack(direction = ROW, padding_top = 5))
        btnAddExcel_wndAddCard = toga.Button("📂", on_press = self.cbAddWordList)
        self.txtExcelFilePath_wndAddCard = toga.TextInput(style = Pack(flex = 1), readonly = True)
        boxRow3_wndAddWordList.add(
            btnAddExcel_wndAddCard,
            self.txtExcelFilePath_wndAddCard
        )

        sbMain_wndAddWordList.content = toga.Box(
            children = [
                boxRow1_wndAddWordList,
                boxRow2_wndAddWordList,
                boxRow3_wndAddWordList
            ], 
            style = Pack(direction = COLUMN)
        )
        self.btnSaveNewWordList_wndAddWordList = toga.Button("", style = Pack(flex = 1), on_press = self.cbSaveNewWordList)

        self.wndAddWordList.content = toga.Box(children = [sbMain_wndAddWordList, self.btnSaveNewWordList_wndAddWordList], style = Pack(direction = COLUMN))
        
        # Hidden button
        self.btnStrange = toga.Button("", style = Pack(flex = 1, background_color = "#FF8989"))
        self.btnVague   = toga.Button("", style = Pack(flex = 1, background_color = "#FBD85D"))
        self.btnFamiliar= toga.Button("", style = Pack(flex = 1, background_color = "#D0F5BE"))
        self.btnDelete  = toga.Button("🗑️", style = Pack(flex = 1, background_color = "#A1C2F1"))
        self.btnShowBack= toga.Button("", style = Pack(flex = 1))
        
        self.btnShowBack.on_press = self.cbShowBack
        self.btnStrange.on_press = partial(self.cbNextCard, 0)
        self.btnVague.on_press = partial(self.cbNextCard, 2)
        self.btnFamiliar.on_press = partial(self.cbNextCard, 4)
        self.btnDelete.on_press = self.cbNoFurtherReviewThisWord

        global g_bIsMobile
        if GetOS() in ['Windows','Linux','MacOS']:
            g_bIsMobile = False

        ## Read language file 
        with open(g_strLangPath, "r") as f:
            self.m_dictLanguages = json.load(f)
        self.cboChangeLang_wndOpt.items = self.m_dictLanguages.keys()
        
        global g_strDefaultLang, g_strOpenedWordListName, g_bOnlyReview

        #Read configuration
        (iX, iY) = self.main_window.position
        (iWidth, iHeight) = self.main_window.size
        if os.path.exists(g_strConfigPath) == True:
            with open(g_strConfigPath, "r") as f:
                dictConfig = json.load(f)
                g_strDefaultLang = dictConfig['lang']
                g_strOpenedWordListName = dictConfig['cur_wordlist']
                g_bOnlyReview = dictConfig['only_review']
                if g_bIsMobile == False:
                    t_iX = dictConfig['X']
                    t_iY = dictConfig['Y']
                    t_iWidth = dictConfig['width']
                    t_iHeight = dictConfig['height']
                    (iResolutionWidth, iResolutionHeight) = pyautogui.size()
                    if 0 <= t_iX <= t_iX + t_iWidth <= iResolutionWidth and\
                        0 <= t_iY <= t_iY + t_iHeight <= iResolutionHeight:
                        iX = t_iX
                        iY = t_iY
                        iWidth = t_iWidth
                        iHeight = t_iHeight
        self.chkIsReviewOnly_wndMain.value = g_bOnlyReview
        self.main_window.position = (iX, iY)
        self.main_window.size = (iWidth, iHeight)

        ## Change the language. When the content in languge ComboBox is changed, it will jump to the first added language. 
        self.ChangeLanguageAccordingDefaultLang()
        
        # Read wordlists and cards
        if os.path.exists(g_strWordlistsPath) == True:
            with open(g_strWordlistsPath, "r") as f:
                global g_dictWordLists
                g_dictWordLists = json.load(f)
        if os.path.exists(g_strCardsPath) == True:
            with open(g_strCardsPath, "r") as f:
                global g_dictCards
                g_dictCards = json.load(f)
        
        self.wbCard_wndMain.set_content("NO_CONTENT", g_EmptyHtml)

        self.m_lsRememberSeq = []
        
        #Show the main window
        self.main_window.show()
        bAllFilesAreValid = self.ValidateFiles()
        if bAllFilesAreValid == True and os.path.exists(g_strDBPath) == True: # Not support displaying an error dialog when no operation on the just executed program. If so, re-config is required.
            
            lsValidWordlists = []
            conn = sqlite3.connect(g_strDBPath)
            cursor = conn.cursor()
            for eachWordlist in g_dictWordLists:
                cursor.execute("SELECT * FROM statistics WHERE wordlist = ?", (eachWordlist,))
                result = cursor.fetchall()
                if len(result) == 0:
                    continue
                lsValidWordlists.append(eachWordlist)
            conn.commit()
            conn.close()

            if g_strOpenedWordListName not in lsValidWordlists:
                g_strOpenedWordListName = lsValidWordlists[0]
            
            self.cboCurWordlist_wndMain.items = lsValidWordlists
            self.cboCurWordlist_wndMain.value = g_strOpenedWordListName

            self.GenerateNewWordsSeqOfCurrentWordList()
            self.cbNextCard(-1, None)
        
      
    # Callback function  
    def cbOpenWndOpt(self, widget):
        global g_bSubFormReleased
        if g_bSubFormReleased == False and g_bIsMobile == False:
            return
        self.cboChangeLang_wndOpt.value = g_strDefaultLang
        g_bSubFormReleased = False
        self.wndOpt.show()

    def cbCloseWndOpt(self, widget):
        global g_bSubFormReleased
        g_bSubFormReleased = True
        self.SaveConfiguration()
        self.wndOpt.hide()

    def cbOpenWndMgmt(self, widget):
        global g_bSubFormReleased
        if g_bSubFormReleased == False and g_bIsMobile == False: #On the mobile, on_close will not execute
            return
        g_bSubFormReleased = False

        #In case of not saving!!!
        self.m_dictTmpCards = copy.deepcopy(g_dictCards)
        self.m_dictTmpWordLists = copy.deepcopy(g_dictWordLists)

        self.FreshSbCards()
        self.FreshSbWordLists()
        self.wndMgmt.show()

    def cbCloseWndMgmt(self, widget):
        global g_bSubFormReleased
        g_bSubFormReleased = True
        self.wndMgmt.hide()

    def cbChangeLanguageForCboOnSelect(self, widget):
        global g_strDefaultLang
        g_strDefaultLang = widget.value
        self.ChangeLanguageAccordingDefaultLang()
    
    def cbChangeWordListEveryNCards(self, widget):
        global g_iChangeWordListEveryNCards
        g_iChangeWordListEveryNCards = widget.value
    
    def cbLearnBackEveryNWords(self, widget):
        global g_iLearnBackEveryNWords
        g_iLearnBackEveryNWords = widget.value
    
    def cbGetFilePath(self, strFilePath, widget):
        return strFilePath

    def cbAddWordList(self, widget):
        lsFilePaths = []
        self.wndMgmt.open_file_dialog(title = g_strMsgBox_OpenFileDialogTitle, file_types = ["xlsx"],multiselect = False, \
                                      on_result = lambda _wnd, _pathFilePath: lsFilePaths.append(_pathFilePath))
        if lsFilePaths[0] == None:
            return
        strFilePath = lsFilePaths[0].absolute() #Pathlic.WindowsPath object
        try:
            xlsx = load_workbook(strFilePath, read_only = True)
            sheet = xlsx.worksheets[0]
            xlsx.close()
            self.txtExcelFilePath_wndAddCard.value = strFilePath
        except Exception:
            self.wndMgmt.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_InvalidFile)
        
    def cbOpenWndAddCard(self, widget):
        global g_bSubSubFormReleased
        if g_bSubSubFormReleased == False and g_bIsMobile == False:
            return
        g_bSubSubFormReleased = False
        self.wndAddCard.show()

    def cbCloseWndAddCard(self, widget):
        global g_bSubSubFormReleased
        g_bSubSubFormReleased = True
        self.txtFrontFrontendFilePath_wndAddCard.value = ""
        self.txtFrontBackendFilePath_wndAddCard.value = ""
        self.txtBackFrontendFilePath_wndAddCard.value = ""
        self.txtBackBackendFilePath_wndAddCard.value = ""
        self.txtSpecifyNewCardName_wndAddCard.value = ""
        self.wndAddCard.hide()

    def cbEnableBack(self, widget):
        if widget.value == True:
            self.boxRow3_wndAddCard.style.update(visibility = VISIBLE)
        else:
            self.boxRow3_wndAddCard.style.update(visibility = HIDDEN)

    def cbFrontFrontend(self, widget):
        lsFilePaths = []
        self.wndMgmt.open_file_dialog(title = g_strMsgBox_OpenFileDialogTitle, file_types = ["htm", "html"],multiselect = False, \
                                      on_result = lambda _wnd, _pathFilePath: lsFilePaths.append(_pathFilePath))
        if lsFilePaths[0] == None:
            return
        strFilePath = lsFilePaths[0].absolute() #Pathlic.WindowsPath object
        try:
            with open(strFilePath, 'r') as f:
                soup = BeautifulSoup(f.read(), 'html.parser')
            self.txtFrontFrontendFilePath_wndAddCard.value = strFilePath
        except Exception:
            self.wndMgmt.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_InvalidFile)
    
    def cbFrontBackend(self, widget):
        lsFilePaths = []
        self.wndMgmt.open_file_dialog(title = g_strMsgBox_OpenFileDialogTitle, file_types = ["py"],multiselect = False, \
                                      on_result = lambda _wnd, _pathFilePath: lsFilePaths.append(_pathFilePath))
        if lsFilePaths[0] == None:
            return
        strFilePath = lsFilePaths[0].absolute() #Pathlic.WindowsPath object
        try:
            with open(strFilePath, 'r') as f:
                strCode = f.read()
                compile(strCode, strFilePath, 'exec')
            self.txtFrontBackendFilePath_wndAddCard.value = strFilePath
        except Exception:
            self.wndMgmt.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_InvalidFile)
            
    def cbBackFrontend(self, widget):
        lsFilePaths = []
        self.wndMgmt.open_file_dialog(title = g_strMsgBox_OpenFileDialogTitle, file_types = ["htm", "html"],multiselect = False, \
                                      on_result = lambda _wnd, _pathFilePath: lsFilePaths.append(_pathFilePath))
        if lsFilePaths[0] == None:
            return
        strFilePath = lsFilePaths[0].absolute() #Pathlic.WindowsPath object
        try:
            with open(strFilePath, 'r') as f:
                soup = BeautifulSoup(f.read(), 'html.parser')
            self.txtBackFrontendFilePath_wndAddCard.value = strFilePath
        except Exception:
            self.wndMgmt.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_InvalidFile)
    
    def cbBackBackend(self, widget):
        lsFilePaths = []
        self.wndMgmt.open_file_dialog(title = g_strMsgBox_OpenFileDialogTitle, file_types = ["py"],multiselect = False, \
                                      on_result = lambda _wnd, _pathFilePath: lsFilePaths.append(_pathFilePath))
        if lsFilePaths[0] == None:
            return
        strFilePath = lsFilePaths[0].absolute() #Pathlic.WindowsPath object
        try:
            with open(strFilePath, 'r') as f:
                strCode = f.read()
                compile(strCode, strFilePath, 'exec')
            self.txtBackBackendFilePath_wndAddCard.value = strFilePath
        except Exception:
            self.wndMgmt.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_InvalidFile)
    
    def cbValidateNameTxt(self, widget):
        global g_strTmpNameTxt
        if widget.value == "":
            g_strTmpNameTxt = ""
        if not re.match(r"^[a-zA-Z_][a-zA-Z0-9_]*$", widget.value):
            widget.value = g_strTmpNameTxt
            #matches = re.findall(r"[a-zA-Z]+", widget.value)
            #widget.value = ''.join(matches)
        if len(widget.value) > 10:
            widget.value = g_strTmpNameTxt
        g_strTmpNameTxt = widget.value
    
    def cbSaveNewCard(self, widget):
        strCardCname = self.txtSpecifyNewCardName_wndAddCard.value
        if strCardCname == "":
            self.wndAddCard.info_dialog(title = g_strMsgBox_InfoDialogTitle, message = g_strMsgBox_InfoDialogMsg_EmptyCardName)
            return
        if strCardCname in self.m_dictTmpCards:
            self.wndAddCard.info_dialog(title = g_strMsgBox_InfoDialogTitle, message = g_strMsgBox_InfoDialogMsg_RepetitiveCardName)
            return
        
        strFrontFrontendFilePath = self.txtFrontFrontendFilePath_wndAddCard.value
        strFrontBackendFilePath = self.txtFrontBackendFilePath_wndAddCard.value
        strBackFrontendFilePath = self.txtBackFrontendFilePath_wndAddCard.value
        strBackBackendFilePath = self.txtBackBackendFilePath_wndAddCard.value
        bEnableBack = self.chkEnableBack_wndAddCard.value

        if bEnableBack == False:
            if strFrontFrontendFilePath == "" or strFrontBackendFilePath == "":
                self.wndAddCard.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_CannotCreateCard)
                return
            else:
                self.m_dictTmpCards[strCardCname] = {
                    "single-sided": {
                        "frontend": strFrontFrontendFilePath,
                        "backend": strFrontBackendFilePath
                    }
                }
        else:
            if strFrontFrontendFilePath == "" or strFrontBackendFilePath == "" or strBackFrontendFilePath == "" or strBackBackendFilePath == "":
                self.wndAddCard.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_CannotCreateCard)
                return
            else:
                self.m_dictTmpCards[strCardCname] = {
                    "front": {
                        "frontend": strFrontFrontendFilePath,
                        "backend": strFrontBackendFilePath
                    },
                    "back": {
                        "frontend": strBackFrontendFilePath,
                        "backend": strBackBackendFilePath
                    }
                }
        self.cbCloseWndAddCard(None)
        self.FreshSbCards()
        self.FreshSbWordLists()
            

    def cbOpenWndAddWordList(self, widget):
        global g_bSubSubFormReleased
        if g_bSubSubFormReleased == False and g_bIsMobile == False:
            return
        g_bSubSubFormReleased = False
        self.wndAddWordList.show()

    def cbCloseWndAddWordlist(self, widget):
        global g_bSubSubFormReleased
        g_bSubSubFormReleased = True
        self.txtSpecifyNewWordListName_wndAddWordList.value =""
        self.txtExcelFilePath_wndAddCard.value = ""
        self.wndAddWordList.hide()

    def cbSaveNewWordList(self, widget):
        strWordListCname = self.txtSpecifyNewWordListName_wndAddWordList.value
        if strWordListCname == "":
            self.wndAddWordList.info_dialog(title = g_strMsgBox_InfoDialogTitle, message = g_strMsgBox_InfoDialogMsg_EmptyWordListName)
            return
        if strWordListCname in self.m_dictTmpWordLists:
            self.wndAddWordList.info_dialog(title = g_strMsgBox_InfoDialogTitle, message = g_strMsgBox_InfoDialogMsg_RepetitiveWordListName)
            return
        strExcelFilePath = self.txtExcelFilePath_wndAddCard.value
        if strExcelFilePath == "":
            self.wndAddWordList.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_CannotCreateWordlist)
            return
        else:
            self.m_dictTmpWordLists[strWordListCname] = {
                "FilePath": strExcelFilePath,
                "CardType1": "",
                "CardType2": "",
                "NewWordsPerGroup": 20,
                "OldWordsPerGroup": 20,
                "policy": 0
            }
        self.cbCloseWndAddWordlist(None)
        self.FreshSbCards()
        self.FreshSbWordLists()
    
    def cbDeleteCard(self, strCardName, widget):
        setAllUsedCardNames = set()
        for eachWordList in self.m_dictTmpWordLists:
            strWordListCardType1 = self.m_dictTmpWordLists[eachWordList]['CardType1']
            if strWordListCardType1 != "":
                setAllUsedCardNames.add(strWordListCardType1)
            strWordListCardType2 = self.m_dictTmpWordLists[eachWordList]['CardType2']
            if strWordListCardType2 != "":
                setAllUsedCardNames.add(strWordListCardType2) 
        if strCardName in setAllUsedCardNames:
            self.wndMgmt.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_CardHasAssociatedWordlist)
            return
        self.m_dictTmpCards.pop(strCardName)
        self.FreshSbCards()
        self.FreshSbWordLists()

    def cbDeleteWordList(self, strWordListName, widget):
        # If a word list is deleted, the record will be removed.
        lsQuestionDialogResult = []
        self.wndMgmt.question_dialog(title = g_strMsgBox_QuestionDialogTitle_Warning, message = g_strMsgBox_QuestionDialogMsg_WarningOnDeletingWordList,\
                                     on_result = lambda _wnd, _bQuestionResult: lsQuestionDialogResult.append(_bQuestionResult))
        if lsQuestionDialogResult[0] == None or lsQuestionDialogResult[0] == False:
            return
        
        self.m_dictTmpWordLists.pop(strWordListName)
        self.FreshSbCards()
        self.FreshSbWordLists()
    
    def cbChangeWordListCardType(self, strWordList, iDeleteCardType, widget):
        if iDeleteCardType == 1:
            self.m_dictTmpWordLists[strWordList]['CardType1'] = widget.value
        elif iDeleteCardType == 2:
            self.m_dictTmpWordLists[strWordList]['CardType2'] = widget.value

    def cbChangeNewWordsPerGroup(self, strWordList, widget):
        self.m_dictTmpWordLists[strWordList]['NewWordsPerGroup'] = int(widget.value)

    def cbChangeOldWordsPerGroup(self, strWordList, widget):
        self.m_dictTmpWordLists[strWordList]['OldWordsPerGroup'] = int(widget.value)
    
    '''
    def cbChangeReviewPolicy(self, strWordList, widget):
        self.m_dictTmpWordLists[strWordList]['policy'] = widget.value.val
    '''
    def cbChangeReviewPolicy(self, strWordList, widget):
        self.m_dictTmpWordLists[strWordList]['policy'] = widget.items.index(widget.value)
        self.SaveConfiguration()

    def cbApplyChangesToDB(self, widget):
        try:
            # Validate cards
            for eachCard in self.m_dictTmpCards:
                if len(self.m_dictTmpCards[eachCard]) == 1:
                    strHtmlPath = self.m_dictTmpCards[eachCard]['single-sided']['frontend']
                    with open(strHtmlPath, 'r') as f:
                        soup = BeautifulSoup(f.read(), 'html.parser')
                    strPyPath = self.m_dictTmpCards[eachCard]['single-sided']['backend']
                    with open(strPyPath, 'r') as f:
                        strCode = f.read()
                        compile(strCode, strPyPath, 'exec')
                else:

                    strFHtmlPath = self.m_dictTmpCards[eachCard]['front']['frontend']
                    with open(strFHtmlPath, 'r') as f:
                        soup = BeautifulSoup(f.read(), 'html.parser')
                    strFPyPath = self.m_dictTmpCards[eachCard]['front']['backend']
                    with open(strFPyPath, 'r') as f:
                        strCode = f.read()
                        compile(strCode, strFPyPath, 'exec')

                    strBHtmlPath = self.m_dictTmpCards[eachCard]['back']['frontend']
                    with open(strBHtmlPath, 'r') as f:
                        soup = BeautifulSoup(f.read(), 'html.parser')
                    strBPyPath = self.m_dictTmpCards[eachCard]['back']['backend']
                    with open(strBPyPath, 'r') as f:
                        strCode = f.read()
                        compile(strCode, strBPyPath, 'exec')

            for eachWordlist in self.m_dictTmpWordLists:  
            # Validate wordlists
                strXlsxPath = self.m_dictTmpWordLists[eachWordlist]['FilePath'] 
                xlsx = load_workbook(strXlsxPath, read_only = True)
                sheet = xlsx.worksheets[0]
                xlsx.close()
        except Exception:
            self.wndMgmt.error_dialog(title = g_strMsgBox_ErrorDialogTitle, message = g_strMsgBox_ErrorDialogMsg_InvalidFiles)

        if g_bSubSubFormReleased == False and g_bIsMobile == False: #If there is a sub-window opening
            return
        
        # Check if there is no word list
        if len(self.m_dictTmpWordLists) == 0:
            self.wndMgmt.error_dialog(g_strMsgBox_ErrorDialogTitle, g_strMsgBox_ErrorDialogMsg_NoWordList)
            return
        
        # Check all word list is available
        for eachWordList in self.m_dictTmpWordLists:
            if self.m_dictTmpWordLists[eachWordList]['CardType1'] == "":
                self.wndMgmt.error_dialog(g_strMsgBox_ErrorDialogTitle, g_strMsgBox_ErrorDialogMsg_Card1IsNull)
                return
            else:
                if self.m_dictTmpWordLists[eachWordList]['CardType2'] == self.m_dictTmpWordLists[eachWordList]['CardType1']:
                    self.wndMgmt.error_dialog(g_strMsgBox_ErrorDialogTitle, g_strMsgBox_ErrorDialogMsg_Card1SameAsCard2)
                    return
        
        conn = sqlite3.connect(g_strDBPath)
        cursor = conn.cursor()

        #If db is not existed,
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS statistics(
                wordlist        TEXT,
                word_no         INTEGER,
                card_type       INTEGER,
                review_date     DATE,
                easiness        REAL,
                interval        INTEGER,
                repetitions     INTEGER
            )
        ''')

        # If word list in db is not existed in current setting, remove
        cursor.execute("SELECT DISTINCT wordlist FROM statistics")
        lsAllWordListsInDB = [elem[0] for elem in cursor.fetchall()]
        lsAllWordListsInDBRemained = lsAllWordListsInDB[:]
        for eachWordListInDB in lsAllWordListsInDB:
            if eachWordListInDB not in self.m_dictTmpWordLists:
                # In case of saving the record
                lsQuestionDialogResult = []
                self.wndMgmt.question_dialog(title = g_strMsgBox_QuestionDialogTitle_Warning, message = g_strMsgBox_QuestionDialogMsg_NotFoundWordList % eachWordListInDB,\
                                            on_result = lambda _wnd, _bQuestionResult: lsQuestionDialogResult.append(_bQuestionResult))
                if lsQuestionDialogResult[0] == None or lsQuestionDialogResult[0] == False:
                    continue
                cursor.execute(f"DELETE FROM statistics WHERE wordlist='{eachWordListInDB}'")  
                lsAllWordListsInDBRemained.remove(eachWordListInDB)
        
        for eachWordList in self.m_dictTmpWordLists:
            xlsx = load_workbook(self.m_dictTmpWordLists[eachWordList]['FilePath'], read_only = True)
            sheet = xlsx.worksheets[0]
            iSheetMaxRow = sheet.max_row
            xlsx.close()
            if eachWordList in lsAllWordListsInDBRemained:
                # If the word list is in db,  detect number change and update card
                cursor.execute(f"SELECT MAX(word_no) FROM statistics WHERE wordlist=?",(eachWordList,)) 
                iWordListMaxRowInDB = cursor.fetchone()[0]
               
                if self.m_dictTmpWordLists[eachWordList]['CardType2'] == "":    # If the word list does not has card2 but database has, delete
                    cursor.execute("DELETE FROM statistics WHERE wordlist=? AND card_type=?",(eachWordList, 2))
                else:                                                   # If the word list has card2 but database does not have, new
                    cursor.execute('SELECT COUNT(*) FROM statistics WHERE wordlist=? AND card_type=?', (eachWordList, 2))
                    iCntWordListCardType2InDB = cursor.fetchone()[0]
                    if iCntWordListCardType2InDB == 0:
                        for i in range(2, iWordListMaxRowInDB + 1):
                            cursor.execute("INSERT INTO statistics (wordlist, word_no, card_type) VALUES (?, ?, ?)", (eachWordList, i, 2))
                
                if iWordListMaxRowInDB < iSheetMaxRow:
                    for i in range(iWordListMaxRowInDB + 1, iSheetMaxRow + 1):
                        cursor.execute("INSERT INTO statistics (wordlist, word_no, card_type) VALUES (?, ?, ?)", (eachWordList, i, 1))
                        if self.m_dictTmpWordLists[eachWordList]['CardType2'] != "":
                            cursor.execute("INSERT INTO statistics (wordlist, word_no, card_type) VALUES (?, ?, ?)", (eachWordList, i, 2))
                elif iWordListMaxRowInDB > iSheetMaxRow:
                    #
                    cursor.execute('DELETE FROM statistics WHERE wordlist=? AND word_no BETWEEN ? AND ?', (eachWordList ,iSheetMaxRow + 1, iWordListMaxRowInDB))
                    
            else:
                # If the word list is not in db, add directly
                for i in range(2, iSheetMaxRow + 1):
                    cursor.execute("INSERT INTO statistics (wordlist, word_no, card_type) VALUES (?, ?, ?)", (eachWordList, i, 1))
                    if self.m_dictTmpWordLists[eachWordList]['CardType2'] != "":
                        cursor.execute("INSERT INTO statistics (wordlist, word_no, card_type) VALUES (?, ?, ?)", (eachWordList, i, 2))
        conn.commit()
        conn.close()
        
        global g_dictWordLists, g_dictCards # Apply 
        g_dictWordLists = copy.deepcopy(self.m_dictTmpWordLists)
        g_dictCards = copy.deepcopy(self.m_dictTmpCards)

        self.m_dictTmpWordLists.clear() #Clear temporary
        self.m_dictTmpCards.clear()
        
        self.SaveWordListsAndCardsToFiles()

        self.wndMgmt.info_dialog(title = g_strMsgBox_InfoDialogTitle, message = g_strMsgBox_InfoDialogMsg_Completed)

        lsValidWordlists = []
        conn = sqlite3.connect(g_strDBPath)
        cursor = conn.cursor()
        for eachWordlist in g_dictWordLists:
            cursor.execute("SELECT * FROM statistics WHERE wordlist = ?", (eachWordlist,))
            result = cursor.fetchall()
            if len(result) == 0:
                continue
        lsValidWordlists.append(eachWordlist)
        conn.commit()
        conn.close()
        global g_strOpenedWordListName
        if g_strOpenedWordListName not in lsValidWordlists:
            g_strOpenedWordListName = lsValidWordlists[0]

        self.cboCurWordlist_wndMain.items = lsValidWordlists
        self.cboCurWordlist_wndMain.value = g_strOpenedWordListName

        self.GenerateNewWordsSeqOfCurrentWordList()
        self.cbNextCard(-1, None)

        self.cbCloseWndMgmt(None)

    def cbChangeChkIsReviewOnly(self, widget):
        global g_bOnlyReview
        g_bOnlyReview = widget.value
        if g_bOnlyReview == False:
            self.cbNextCard(-1, None)
            

    def cbNoFurtherReviewThisWord(self, widget):
        lsQuestionDialogResult = []
        self.main_window.question_dialog(title = g_strMsgBox_QuestionDialogTitle_Warning, message = g_strMsgBox_QuestionDialogMsg_DeleteThisWord,\
                                    on_result = lambda _wnd, _bQuestionResult: lsQuestionDialogResult.append(_bQuestionResult))
        if lsQuestionDialogResult[0] == None or lsQuestionDialogResult[0] == False:
            return
        self.cbNextCard(5, None)

    def cbNextCard(self, iQuality, widget): # iQuality = -1 if there is no need to update record
        self.ChangeWordLearningBtns(0)
        conn = sqlite3.connect(g_strDBPath)
        cursor = conn.cursor()

        global g_bSwitchContinueNewWords, g_iTmpCntStudiedNewWords, g_iTmpCntStudiedOldWords, g_iTmpPrevWordNo, g_iTmpPrevCardType

        today = datetime.today().date()
        strToday = datetime.now().strftime('%Y-%m-%d')

        # Write old word record
        if iQuality != -1:
            cursor.execute('SELECT * FROM statistics WHERE wordlist=? AND word_no=? AND card_type=? AND review_date IS NULL', \
                           (g_strOpenedWordListName, g_iTmpPrevWordNo, g_iTmpPrevCardType))
            result = cursor.fetchall()
            if len(result) != 0:
                #It is a new word
                review = SMTwo.first_review(iQuality, strToday)

                sql = """
                    UPDATE statistics 
                    SET review_date = ?, easiness = ?, interval = ?, repetitions = ? 
                    WHERE wordlist = ? AND word_no = ? AND card_type = ?
                """
                if iQuality == 0:
                    cursor.execute(sql, (today, review.easiness, review.interval, review.repetitions, g_strOpenedWordListName, g_iTmpPrevWordNo, g_iTmpPrevCardType))
                else:
                    cursor.execute(sql, (review.review_date, review.easiness, review.interval, review.repetitions, g_strOpenedWordListName, g_iTmpPrevWordNo, g_iTmpPrevCardType))
                
            else:
                sql = '''
                    SELECT easiness, interval, repetitions
                    FROM statistics
                    WHERE wordlist = ? AND word_no = ? AND card_type = ?
                '''
                cursor.execute(sql, (g_strOpenedWordListName, g_iTmpPrevWordNo, g_iTmpPrevCardType))
                result = cursor.fetchall()
                easiness, interval, repetitions = result[0]   
                review = SMTwo(easiness, interval, repetitions).review(iQuality, strToday)
                sql = """
                    UPDATE statistics 
                    SET review_date = ?, easiness = ?, interval = ?, repetitions = ? 
                    WHERE wordlist = ? AND word_no = ? AND card_type = ?
                """
                if iQuality == 0:
                    cursor.execute(sql, (today, review.easiness, review.interval, review.repetitions, g_strOpenedWordListName, g_iTmpPrevWordNo, g_iTmpPrevCardType))
                else:
                    cursor.execute(sql, (review.review_date, review.easiness, review.interval, review.repetitions, g_strOpenedWordListName, g_iTmpPrevWordNo, g_iTmpPrevCardType))
        if iQuality == 5:
            # Delete card = Use an huge date
            sql = """
                UPDATE statistics
                SET review_date = '9999-12-31'
                WHERE wordlist = ? AND word_no = ? AND card_type = ?
            """
            cursor.execute(sql,(g_strOpenedWordListName, g_iTmpPrevWordNo, g_iTmpPrevCardType))
        # Update tag
        cursor.execute("SELECT COUNT(*) FROM statistics WHERE wordlist = ? AND review_date IS NULL",(g_strOpenedWordListName,))
        iCntStudiedCards = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM statistics WHERE wordlist = ? AND review_date IS NOT NULL",(g_strOpenedWordListName,))
        iCntRestNewCards = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM statistics WHERE wordlist = ? AND review_date <= ?", (g_strOpenedWordListName, strToday))
        iCntDueTimeCards = cursor.fetchone()[0]

        self.lblLearningProgressValue_wndMain.text = str(iCntStudiedCards)+"("+ str(g_iTmpCntStudiedNewWords) + "," + str(g_iTmpCntStudiedOldWords) \
            + ")/" +str(iCntRestNewCards) + "/"+str(iCntDueTimeCards)

        bFindAnOldWord = False
        # Update the learning process text
        if g_bOnlyReview == False:
            if iCntDueTimeCards != 0:
                if iCntRestNewCards != 0:
                    # Get if there is no card needing reviewing
                    if g_dictWordLists[g_strOpenedWordListName]['policy'] == 0:
                        # Learn new words first
                        if g_iTmpCntStudiedNewWords == 0:
                                bFindAnOldWord = False
                                g_bSwitchContinueNewWords = True
                        else:
                            if (g_iTmpCntStudiedNewWords % g_dictWordLists[g_strOpenedWordListName]['NewWordsPerGroup'] == 0)\
                                and g_bSwitchContinueNewWords == True:
                                bFindAnOldWord = False
                                g_bSwitchContinueNewWords = False
                            elif g_iTmpCntStudiedOldWords % g_dictWordLists[g_strOpenedWordListName]['OldWordsPerGroup'] == 0 \
                                and g_bSwitchContinueNewWords == False:
                                bFindAnOldWord = True
                                g_bSwitchContinueNewWords = True
                            else:
                                bFindAnOldWord = not g_bSwitchContinueNewWords
                    elif g_dictWordLists[g_strOpenedWordListName]['policy'] == 1:
                        # Randomly
                        iRandom = random.randint(1, g_dictWordLists[g_strOpenedWordListName]['NewWordsPerGroup'] + g_dictWordLists[g_strOpenedWordListName]['OldWordsPerGroup'])
                        if iRandom < g_dictWordLists[g_strOpenedWordListName]['NewWordsPerGroup']:
                            bFindAnOldWord = True
                        else:
                            bFindAnOldWord = False
                    elif g_dictWordLists[g_strOpenedWordListName]['policy'] == 2:
                        # Review first
                        if g_iTmpCntStudiedOldWords == 0:
                            bFindAnOldWord = True
                            g_bSwitchContinueNewWords = False
                        else:
                            if (g_iTmpCntStudiedOldWords % g_dictWordLists[g_strOpenedWordListName]['OldWordsPerGroup'] == 0)\
                                and g_bSwitchContinueNewWords == False:
                                bFindAnOldWord == True
                                g_bSwitchContinueNewWords = True
                            elif (g_iTmpCntStudiedNewWords % g_dictWordLists[g_strOpenedWordListName]['NewWordsPerGroup'] == 0)\
                                and g_bSwitchContinueNewWords == True:
                                bFindAnOldWord = False
                                g_bSwitchContinueNewWords = False
                            else:
                                bFindAnOldWord = not g_bSwitchContinueNewWords
                else: # There is no new cards
                    bFindAnOldWord = True
            else: # All card has been reviewed
                bFindAnOldWord = False
        else:
            bFindAnOldWord = True
        
        if bFindAnOldWord == True:
            conn.create_function('power', 2, lambda _x, _y: _x**_y) #SQLite does not support exponentiation 
            sql = """
                SELECT word_no, card_type 
                FROM statistics
                WHERE wordlist = ? 
                    AND review_date IS NOT NULL 
                    AND review_date <= ?
                ORDER BY review_date ASC, ( power(easiness, repetitions) / interval) ASC
                LIMIT 1;
            """
            cursor.execute(sql, (g_strOpenedWordListName, strToday))
            result = cursor.fetchone()
            
            if result == None:
                self.wbCard_wndMain.set_content("NO_CONTENT", g_EmptyHtml)
                conn.commit()
                conn.close()
                # If only review is true and rest new cards is not 0,
                if g_bOnlyReview == True and iCntRestNewCards > 0:
                    if g_iTmpCntStudiedNewWords == 0 and g_iTmpCntStudiedOldWords == 0: # A bug: if the program just runs, the info dialog will raise RuntimeError.
                        return
                    self.main_window.info_dialog(g_strMsgBox_InfoDialogTitle, g_strMsgBox_InfoDialogMsg_ContinueToStudyNewWords)
                    return
                # All old words have been studied.
                if g_iTmpCntStudiedNewWords == 0 and g_iTmpCntStudiedOldWords == 0: # A bug: if the program just runs, the info dialog will raise RuntimeError.
                    return
                self.main_window.info_dialog(g_strMsgBox_InfoDialogTitle, g_strMsgBox_InfoDialogMsg_MissionComplete)
                return
            else:
                g_iTmpPrevWordNo = result[0]
                g_iTmpPrevCardType = result[1]
                g_iTmpCntStudiedOldWords += 1
            
            sql = '''
                SELECT easiness, interval, repetitions
                FROM statistics
                WHERE wordlist = ? AND word_no = ? AND card_type = ?
            '''
            cursor.execute(sql, (g_strOpenedWordListName, g_iTmpPrevWordNo, g_iTmpPrevCardType))
            result = cursor.fetchall()
            easiness, interval, repetitions = result[0]

            review = SMTwo(easiness, interval, repetitions).review(0, strToday)
            self.btnStrange.text = "😞 (0 d)"

            review2 = SMTwo(easiness, interval, repetitions).review(2, strToday)
            self.btnVague.text = "😐 ("+str((review2.review_date - today).days)+" d)"

            review3 = SMTwo(easiness, interval, repetitions).review(4, strToday)
            self.btnFamiliar.text = "😄 ("+str((review3.review_date - today).days)+" d)"

        else:
            #Find a new word
            if len(self.m_lsRememberSeq) == 0:
                # All new words have been studied. 
                
                self.wbCard_wndMain.set_content("NO_CONTENT", g_EmptyHtml)
                conn.commit()
                conn.close()
                if g_iTmpCntStudiedNewWords == 0 and g_iTmpCntStudiedOldWords == 0: # A bug: if the program just runs, the info dialog will raise RuntimeError.
                    return
                self.main_window.info_dialog(g_strMsgBox_InfoDialogTitle, g_strMsgBox_InfoDialogMsg_MissionComplete)
                return
            else:
                g_iTmpPrevWordNo = self.m_lsRememberSeq[0][0]
                g_iTmpPrevCardType = self.m_lsRememberSeq[0][1]
                g_iTmpCntStudiedNewWords += 1
                self.m_lsRememberSeq.pop(0)
            
            review = SMTwo.first_review(0, strToday)
            self.btnStrange.text = "😞 (0 d)"
            
            review2 = SMTwo.first_review(2, strToday)
            self.btnVague.text = "😐 ("+str((review2.review_date - today).days)+" d)"
            
            review3 = SMTwo.first_review(4, strToday)
            self.btnFamiliar.text = "😄 ("+str((review3.review_date - today).days)+" d)"
        
        conn.commit()
        conn.close()

        dictCardTypeThisCardUses = g_dictCards[g_dictWordLists[g_strOpenedWordListName]['CardType'+str(g_iTmpPrevCardType)]]
        if len(dictCardTypeThisCardUses) == 1:
            #Display web
            xlsx = load_workbook(g_dictWordLists[g_strOpenedWordListName]['FilePath'])
            sheet = xlsx.worksheets[0]
            dicDynVariables = {}
            for i in range(1, sheet.max_column + 1):
                dicDynVariables[ConvertNumToExcelColTitle(i)] = FilterForExcelNoneValue(sheet.cell(row = g_iTmpPrevWordNo, column = i).value)
            strRenderedHTML = GetRenderedHTML(
                dictCardTypeThisCardUses['single-sided']['frontend'],
                dictCardTypeThisCardUses['single-sided']['backend'],
                dicDynVariables
            )
            self.wbCard_wndMain.set_content("wordgets", strRenderedHTML)

            #Change button
            self.ChangeWordLearningBtns(2)
        else:
            #Display web
            
            xlsx = load_workbook(g_dictWordLists[g_strOpenedWordListName]['FilePath'])
            sheet = xlsx.worksheets[0]
            dicDynVariables = {}
            for i in range(1, sheet.max_column + 1):
                dicDynVariables[ConvertNumToExcelColTitle(i)] = FilterForExcelNoneValue(sheet.cell(row = g_iTmpPrevWordNo, column = i).value)
            strRenderedHTML = GetRenderedHTML(
                dictCardTypeThisCardUses['front']['frontend'],
                dictCardTypeThisCardUses['front']['backend'],
                dicDynVariables
            )
            self.wbCard_wndMain.set_content("wordgets", strRenderedHTML)

            #Change button
            self.ChangeWordLearningBtns(1)
        
    def cbShowBack(self, widget):
        self.ChangeWordLearningBtns(0)
        dictCardTypeThisCardUses = g_dictCards[g_dictWordLists[g_strOpenedWordListName]['CardType'+str(g_iTmpPrevCardType)]]

        xlsx = load_workbook(g_dictWordLists[g_strOpenedWordListName]['FilePath'])
        sheet = xlsx.worksheets[0]
        dicDynVariables = {}
        for i in range(1, sheet.max_column + 1):
            dicDynVariables[ConvertNumToExcelColTitle(i)] = FilterForExcelNoneValue(sheet.cell(row = g_iTmpPrevWordNo, column = i).value)
        strRenderedHTML = GetRenderedHTML(
            dictCardTypeThisCardUses['back']['frontend'],
            dictCardTypeThisCardUses['back']['backend'],
            dicDynVariables
        )

        self.wbCard_wndMain.set_content("wordgets", strRenderedHTML)

        self.ChangeWordLearningBtns(2)
   
    def ValidateFiles(self):
        if len(g_dictCards) == 0 or len(g_dictWordLists) == 0:
            return False
        try:
            # Validate cards
            for eachCard in g_dictCards:
                if len(g_dictCards[eachCard]) == 1:
                    strHtmlPath = g_dictCards[eachCard]['single-sided']['frontend']
                    with open(strHtmlPath, 'r') as f:
                        soup = BeautifulSoup(f.read(), 'html.parser')
                    strPyPath = g_dictCards[eachCard]['single-sided']['backend']
                    with open(strPyPath, 'r') as f:
                        strCode = f.read()
                        compile(strCode, strPyPath, 'exec')
                else:

                    strFHtmlPath = g_dictCards[eachCard]['front']['frontend']
                    with open(strFHtmlPath, 'r') as f:
                        soup = BeautifulSoup(f.read(), 'html.parser')
                    strFPyPath = g_dictCards[eachCard]['front']['backend']
                    with open(strFPyPath, 'r') as f:
                        strCode = f.read()
                        compile(strCode, strFPyPath, 'exec')

                    strBHtmlPath = g_dictCards[eachCard]['back']['frontend']
                    with open(strBHtmlPath, 'r') as f:
                        soup = BeautifulSoup(f.read(), 'html.parser')
                    strBPyPath = g_dictCards[eachCard]['back']['backend']
                    with open(strBPyPath, 'r') as f:
                        strCode = f.read()
                        compile(strCode, strBPyPath, 'exec')

            for eachWordlist in g_dictWordLists:  
                # Validate wordlists
                strXlsxPath = g_dictWordLists[eachWordlist]['FilePath'] 
                xlsx = load_workbook(strXlsxPath, read_only = True)
                sheet = xlsx.worksheets[0]
                iMaxRowInSheet = sheet.max_row
                xlsx.close()
                
                conn = sqlite3.connect(g_strDBPath)
                cursor = conn.cursor()
                cursor.execute(f"SELECT MAX(word_no) FROM statistics WHERE wordlist=?",(eachWordlist,)) 
                iWMaxRowInDB = cursor.fetchone()[0]
                conn.commit()
                conn.close()
                if iMaxRowInSheet != iWMaxRowInDB:
                    raise FileNotFoundError
            return True
        except Exception:
            return False
            

    def GenerateNewWordsSeqOfCurrentWordList(self):
        conn = sqlite3.connect(g_strDBPath)
        cursor = conn.cursor()
        cursor.execute("SELECT word_no FROM statistics WHERE wordlist = ? AND review_date IS NULL AND card_type = 1",(g_strOpenedWordListName,))
        lsCard_Type1 = [elem[0] for elem in cursor.fetchall()]
        cursor.execute("SELECT word_no FROM statistics WHERE wordlist = ? AND review_date IS NULL AND card_type = 2",(g_strOpenedWordListName,))
        lsCard_Type2 = [elem[0] for elem in cursor.fetchall()]
        conn.commit()
        conn.close()
        
        self.m_lsRememberSeq = []

        # No new words
        if len(lsCard_Type1) == len(lsCard_Type2) == 0:
            return
        
        # Only one card type
        if len(lsCard_Type2) == 0:
            for i in range(len(lsCard_Type1)):
                self.m_lsRememberSeq.append((lsCard_Type1[i],1))

        # Rest type-1 cards are less than rest type-2 cards 
        if len(lsCard_Type1) < len(lsCard_Type2):
            if len(lsCard_Type1) == 0: # type-1 cards are all studied
                for i in range(len(lsCard_Type2)):
                    self.m_lsRememberSeq.append((lsCard_Type2[i],2))
                return
            else: #Otherwise, first remember the extra type-2 cards
                for i in range(len(lsCard_Type2)):
                    if lsCard_Type2[i] != lsCard_Type1[0]:
                        self.m_lsRememberSeq.append((lsCard_Type2[i],2))
                    else:
                        break
        iNewWordsPerGroup = g_dictWordLists[g_strOpenedWordListName]['NewWordsPerGroup']
        lsGroups1 = []
        t = 1
        lsNumbersAGroup = []
        for i in range(len(lsCard_Type1)):
            if t == iNewWordsPerGroup or i == len(lsCard_Type1) - 1:
                lsNumbersAGroup.append([lsCard_Type1[i],1])
                lsGroups1.append(lsNumbersAGroup[:])
                lsNumbersAGroup.clear()
                t = 1
            else:
                lsNumbersAGroup.append([lsCard_Type1[i],1])
                t += 1
        lsGroups2 = copy.deepcopy(lsGroups1)
        for i in range(len(lsGroups2)):
            for j in range(len(lsGroups2[i])):
                lsGroups2[i][j][1] = 2
        result = [item for pair in zip(lsGroups1, lsGroups2) for item in pair]
        for elem in result:
            self.m_lsRememberSeq.extend(elem)
    
    def cbChangeCurWordList(self, widget):
        global g_iTmpPrevWordNo, g_iTmpPrevCardType, g_iTmpCntStudiedNewWords, g_iTmpCntStudiedOldWords, g_strOpenedWordListName
        if widget.value == g_strOpenedWordListName:
            return
        g_strOpenedWordListName = widget.value
        g_iTmpPrevWordNo = -1
        g_iTmpPrevCardType = -1
        g_iTmpCntStudiedNewWords = 0
        g_iTmpCntStudiedOldWords = 0
        self.lblLearningProgressValue_wndMain.text = ""
        self.GenerateNewWordsSeqOfCurrentWordList()
        self.cbNextCard(-1, None)
        self.SaveConfiguration()

    # Function
    def ChangeLanguageAccordingDefaultLang(self):
        self.lblLang_wndOpt.text                        = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_LANG_TEXT']
        self.lblCard_wndMgmt.text                       = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_CARD_TEXT']
        self.lblWordList_wndMgmt.text                   = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_WORDLIST_TEXT']
        self.wndOpt.title                               = self.m_dictLanguages[g_strDefaultLang]['STR_WND_OPT_TITLE']
        self.wndMgmt.title                              = self.m_dictLanguages[g_strDefaultLang]['STR_WND_MGMT_TITLE']
        self.chkIsReviewOnly_wndMain.text               = self.m_dictLanguages[g_strDefaultLang]['STR_CHK_ISREVIEWONLY_TEXT']
        self.btnShowBack.text                           = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_SHOWBACK_TEXT']
        self.wndAddCard.title                           = self.m_dictLanguages[g_strDefaultLang]['STR_WND_ADDCARD_TITLE']
        self.lblFrontFrontend_WndAddCard.text           = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_FRONT_FRONTEND_TEXT']
        self.lblFrontBackend_WndAddCard.text            = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_FRONT_BACKEND_TEXT']
        self.lblBackFrontend_WndAddCard.text            = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_BACK_FRONTEND_TEXT']
        self.lblBackBackend_WndAddCard.text             = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_BACK_BACKEND_TEXT']
        self.chkEnableBack_wndAddCard.text              = self.m_dictLanguages[g_strDefaultLang]['STR_CHK_ENABLEBACK_TEXT']
        self.btnSaveNewCard_wndAddCard.text             = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_SAVENEWCARD_TEXT']
        self.lblSpecifyNewCardName_wndAddCard.text      = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_SPECIFYNEWCARDNAME_TEXT']
        self.wndAddWordList.title                       = self.m_dictLanguages[g_strDefaultLang]['STR_WND_ADDWORDLIST_TITLE']
        self.lblExcelFile_WndAddWordList.text           = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_EXCELFILE_TEXT']
        self.btnSaveNewWordList_wndAddWordList.text     = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_SAVENEWWORDLIST_TEXT']
        self.lblSpecifyNewWordListName_wndAddWordList.text=self.m_dictLanguages[g_strDefaultLang]['STR_LBL_SPECIFYNEWWORDISTNAME_TEXT']
        self.btnApply_wndMgmt.text                      = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_APPLY_TEXT']
        self.txtSpecifyNewWordListName_wndAddWordList.placeholder = self.m_dictLanguages[g_strDefaultLang]['STR_TXT_PLACEHOLDER_LEGALNAME']
        self.txtSpecifyNewCardName_wndAddCard.placeholder = self.m_dictLanguages[g_strDefaultLang]['STR_TXT_PLACEHOLDER_LEGALNAME']
        self.lblLearningProgressItem_wndMain.text       = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_LEARNINGPROGRESSITEM_TEXT']
        self.btnVisitOfficialWebsite_wndOpt.text        = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_VISITOFFICIALWEBSITE_TEXT']
        global g_strMsgBox_OpenFileDialogTitle, g_strMsgBox_ErrorDialogTitle, g_strMsgBox_ErrorDialogMsg_InvalidFile,\
            g_strMsgBox_InfoDialogTitle, g_strMsgBox_InfoDialogMsg_EmptyCardName, g_strMsgBox_InfoDialogMsg_RepetitiveCardName,\
            g_strMsgBox_InfoDialogMsg_EmptyWordListName, g_strMsgBox_InfoDialogMsg_RepetitiveWordListName,\
            g_strMsgBox_ErrorDialogMsg_CannotCreateCard, g_strMsgBox_ErrorDialogMsg_CannotCreateWordlist,\
            g_strMsgBox_ErrorDialogMsg_CardHasAssociatedWordlist, g_strMsgBox_ErrorDialogMsg_Card1IsNull,\
            g_strMsgBox_QuestionDialogTitle_Warning, g_strMsgBox_QuestionDialogMsg_WarningOnDeletingWordList,\
            g_strMsgBox_QuestionDialogMsg_NotFoundWordList, g_strMsgBox_ErrorDialogMsg_NoWordList,\
            g_strMsgBox_ErrorDialogMsg_Card1SameAsCard2, g_strMsgBox_InfoDialogMsg_Completed, g_strMsgBox_QuestionDialogMsg_DeleteThisWord,\
            g_strMsgBox_InfoDialogMsg_ContinueToStudyNewWords, g_strMsgBox_InfoDialogMsg_MissionComplete,g_strMsgBox_ErrorDialogMsg_InvalidFiles
        g_strMsgBox_OpenFileDialogTitle                 = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_OPENFILEDIALOG_TITLE']
        g_strMsgBox_ErrorDialogTitle                    = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_TITLE']
        g_strMsgBox_ErrorDialogMsg_InvalidFile          = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_MSG_INVALIDFILE']
        g_strMsgBox_InfoDialogTitle                     = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_INFODIALOG_TITLE']
        g_strMsgBox_InfoDialogMsg_EmptyCardName         = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_INFODIALOG_MSG_EMPTYCARDNAME']
        g_strMsgBox_InfoDialogMsg_RepetitiveCardName    = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_INFODIALOG_MSG_REPETITIVECARDNAME']
        g_strMsgBox_InfoDialogMsg_EmptyWordListName     = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_INFODIALOG_MSG_EMPTYWORDLISTNAME']
        g_strMsgBox_InfoDialogMsg_RepetitiveWordListName= self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_INFODIALOG_MSG_REPETITIVEWORDLISTNAME']
        g_strMsgBox_ErrorDialogMsg_CannotCreateCard     = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_MSG_CANNOTCREATECARD']
        g_strMsgBox_ErrorDialogMsg_CannotCreateWordlist = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_MSG_CANNOTCREATEWORDLIST']
        g_strMsgBox_ErrorDialogMsg_CardHasAssociatedWordlist = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_MSG_CARDHASASSOCIATEDWORDLIST']
        g_strMsgBox_ErrorDialogMsg_Card1IsNull          = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_MSG_CARD1ISNULL']
        g_strMsgBox_QuestionDialogTitle_Warning         = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_QUESTIONDIALOG_TITLE_WARNING']
        g_strMsgBox_QuestionDialogMsg_WarningOnDeletingWordList = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_QUESTIONDIALOG_MSG_WARNINGONDELETINGWORDLIST']
        g_strMsgBox_QuestionDialogMsg_NotFoundWordList  = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_QUESTIONDIALOG_MSG_NOTFOUNDWORDLIST']
        g_strMsgBox_ErrorDialogMsg_NoWordList           = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_MSG_NOWORDLIST']
        g_strMsgBox_ErrorDialogMsg_Card1SameAsCard2     = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_MSG_CARD1SAMEASCARD2']
        g_strMsgBox_InfoDialogMsg_Completed             = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_INFODIALOG_MSG_COMPLETED']
        g_strMsgBox_QuestionDialogMsg_DeleteThisWord    = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_QUESTIONDIALOG_MSG_DELETETHISWORD']
        g_strMsgBox_InfoDialogMsg_ContinueToStudyNewWords= self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_INFODIALOG_MSG_CONTINUETOSTUDYINGNEWWORDS']
        g_strMsgBox_InfoDialogMsg_MissionComplete       = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_INFODIALOG_MSG_MISSIONCOMPLETE']
        g_strMsgBox_ErrorDialogMsg_InvalidFiles          = self.m_dictLanguages[g_strDefaultLang]['STR_MSGBOX_ERRORDIALOG_MSG_INVALIDFILES']
        global g_strFront, g_strBack, g_strOnesided, g_strFrontend, g_strBackend, g_strFilePath, g_strCard1, g_strCard2
        g_strFront                                      = self.m_dictLanguages[g_strDefaultLang]['STR_FRONT']
        g_strBack                                       = self.m_dictLanguages[g_strDefaultLang]['STR_BACK']
        g_strOnesided                                   = self.m_dictLanguages[g_strDefaultLang]['STR_ONESIDED']
        g_strFrontend                                   = self.m_dictLanguages[g_strDefaultLang]['STR_FRONTEND']
        g_strBackend                                    = self.m_dictLanguages[g_strDefaultLang]['STR_BACKEND']
        g_strFilePath                                   = self.m_dictLanguages[g_strDefaultLang]['STR_FILEPATH']
        g_strCard1                                      = self.m_dictLanguages[g_strDefaultLang]['STR_CARD_1']
        g_strCard2                                      = self.m_dictLanguages[g_strDefaultLang]['STR_CARD_2']
        global g_strLblNewWordsPerGroup, g_strLblOldWordsPerGroup, g_strLblReviewPolicy, g_strCboReviewPolicy_item_LearnFirst, g_strCboReviewPolicy_item_Random, g_strCboReviewPolicy_item_ReviewFirst
        g_strLblNewWordsPerGroup                        = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_NEWWORDSPERGROUP']
        g_strLblOldWordsPerGroup                        = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_OLDWORDSPERGROUP']
        g_strLblReviewPolicy                            = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_REVIEWPOLICY']
        g_strCboReviewPolicy_item_LearnFirst            = self.m_dictLanguages[g_strDefaultLang]['STR_CBO_REVIEWPOLICY_ITEM_LEARNFIRST']
        g_strCboReviewPolicy_item_Random                = self.m_dictLanguages[g_strDefaultLang]['STR_CBO_REVIEWPOLICY_ITEM_RANDOM']
        g_strCboReviewPolicy_item_ReviewFirst           = self.m_dictLanguages[g_strDefaultLang]['STR_CBO_REVIEWPOLICY_ITEM_REVIEWFIRST']


    def FreshSbCards(self):
        while len(self.boxCards_wndMgmt.children) != 0:
            self.boxCards_wndMgmt.remove(self.boxCards_wndMgmt.children[0])
        for eachCard in self.m_dictTmpCards:
            boxCard = toga.Box(style = Pack(direction = ROW))
            mtxtCardInfo = toga.MultilineTextInput(style = Pack(flex = 1), readonly = True)
            strCardInfo = eachCard +"\n"
            if len(self.m_dictTmpCards[eachCard].keys()) == 1:
                strCardInfo += f"\t{g_strOnesided}\n"
                strCardInfo += f"\t\t{g_strFrontend} {g_strFilePath}:"+ self.m_dictTmpCards[eachCard]['single-sided']['frontend'] +"\n"
                strCardInfo += f"\t\t{g_strBackend} {g_strFilePath}:" + self.m_dictTmpCards[eachCard]['single-sided']['backend']
            else:
                strCardInfo += f"\t{g_strFront}\n"
                strCardInfo += f"\t\t{g_strFrontend} {g_strFilePath}:"+ self.m_dictTmpCards[eachCard]['front']['frontend'] +"\n"
                strCardInfo += f"\t\t{g_strBackend} {g_strFilePath}:" + self.m_dictTmpCards[eachCard]['front']['backend'] +"\n"
                
                strCardInfo += f"\t{g_strBack}\n"
                strCardInfo += f"\t\t{g_strFrontend} {g_strFilePath}:"+ self.m_dictTmpCards[eachCard]['back']['frontend'] +"\n"
                strCardInfo += f"\t\t{g_strBackend} {g_strFilePath}:" + self.m_dictTmpCards[eachCard]['back']['backend']
            mtxtCardInfo.value = strCardInfo

            btnDeleteCard = toga.Button("❌", on_press = partial(self.cbDeleteCard, eachCard))
            boxCard.add(
                mtxtCardInfo,
                btnDeleteCard
            )
            self.boxCards_wndMgmt.add(boxCard)

    def FreshSbWordLists(self):
        while len(self.boxWordLists_wndMgmt.children) != 0:
            self.boxWordLists_wndMgmt.remove(self.boxWordLists_wndMgmt.children[0])
        lsAllCardTypes = [""]
        for eachCard in self.m_dictTmpCards:
            lsAllCardTypes.append(eachCard)

        for eachWordList in self.m_dictTmpWordLists:
            boxWordList = toga.Box(style = Pack(direction = COLUMN))

            boxWordListInfo = toga.Box(style = Pack(direction = ROW))
            txtWordListInfo = toga.TextInput(style = Pack(flex = 1), readonly = True)

            strWordListInfo = eachWordList + f" {g_strFilePath}:" + self.m_dictTmpWordLists[eachWordList]['FilePath']
            txtWordListInfo.value = strWordListInfo
            
            btnDeleteWordList = toga.Button("❌", on_press = partial(self.cbDeleteWordList, eachWordList))
            boxWordListInfo.add(
                txtWordListInfo,
                btnDeleteWordList
            )

            boxWordListConfig = toga.Box(style = Pack(direction = COLUMN))

            boxWordListCards = toga.Box(style = Pack(direction = ROW, alignment = CENTER))


            lblCard1 = toga.Label(g_strCard1)
            cboCard1 = toga.Selection(style = Pack(flex = 1), items = lsAllCardTypes, \
                                      on_select = partial(self.cbChangeWordListCardType, eachWordList, 1))
            cboCard1.value = self.m_dictTmpWordLists[eachWordList]["CardType1"]

            lblCard2 = toga.Label(g_strCard2)
            cboCard2 = toga.Selection(style = Pack(flex = 1), items = lsAllCardTypes, \
                                      on_select = partial(self.cbChangeWordListCardType, eachWordList, 2))
            cboCard2.value = self.m_dictTmpWordLists[eachWordList]["CardType2"]

            boxWordListCards.add(
                lblCard1,
                cboCard1,
                lblCard2,
                cboCard2
            )
            
            boxWordListLearningPlan = toga.Box(style = Pack(direction = ROW, alignment = CENTER))

            numNewWordsPerGroup = toga.NumberInput(style = Pack(flex = 1, text_align = RIGHT), min_value = 1, \
                                                   on_change = partial(self.cbChangeNewWordsPerGroup, eachWordList))

            numNewWordsPerGroup.value = self.m_dictTmpWordLists[eachWordList]["NewWordsPerGroup"]

            lblNewWordsPerGroup = toga.Label(g_strLblNewWordsPerGroup)

            numOldWordsPerGroup = toga.NumberInput(style = Pack(flex = 1, padding_left = 5, text_align = RIGHT), min_value = 1,\
                                                   on_change = partial(self.cbChangeOldWordsPerGroup, eachWordList))
            numOldWordsPerGroup.value = self.m_dictTmpWordLists[eachWordList]["OldWordsPerGroup"]

            lblOldWordsPerGroup = toga.Label(g_strLblOldWordsPerGroup)

            lblReviewPolicy = toga.Label(g_strLblReviewPolicy, style = Pack(padding_left = 5))

            '''# This is not supported
            cboReviewPolicy = toga.Selection(style = Pack(flex = 1),\
                                                items = [\
                                                    {"name": g_strCboReviewPolicy_item_LearnFirst,  "val": 0},\
                                                    {"name": g_strCboReviewPolicy_item_Random,      "val": 1},\
                                                    {"name": g_strCboReviewPolicy_item_ReviewFirst, "val": 2},\
                                                ],\
                                                accessor = "name",\
                                                on_select = partial(self.cbChangeReviewPolicy, eachWordList)\
                                            )
            '''
            cboReviewPolicy = toga.Selection(style = Pack(flex = 1),\
                                                items = [g_strCboReviewPolicy_item_LearnFirst, g_strCboReviewPolicy_item_Random, g_strCboReviewPolicy_item_ReviewFirst],\
                                                on_select = partial(self.cbChangeReviewPolicy, eachWordList)\
                                            )
            cboReviewPolicy.value = cboReviewPolicy.items[self.m_dictTmpWordLists[eachWordList]["policy"]]
            boxWordListLearningPlan.add(
                numNewWordsPerGroup,
                lblNewWordsPerGroup,
                numOldWordsPerGroup,
                lblOldWordsPerGroup,
                lblReviewPolicy,
                cboReviewPolicy
            )
            boxWordListConfig.add(
                boxWordListCards,
                boxWordListLearningPlan
            )

            boxWordList.add(
                boxWordListInfo, 
                boxWordListConfig
            )
            self.boxWordLists_wndMgmt.add(boxWordList)

    def ChangeWordLearningBtns(self, iStatus):
        if iStatus == 0: # Clear all buttons
            while len(self.boxRow2Right_wndMain.children) != 0:
                self.boxRow2Right_wndMain.remove(self.boxRow2Right_wndMain.children[0])
            while len(self.boxRow4_wndMain.children) != 0:
                self.boxRow4_wndMain.remove(self.boxRow4_wndMain.children[0])
        elif iStatus == 1: # Show the front
            self.boxRow4_wndMain.add(
                self.btnShowBack
            )
        elif iStatus == 2: # Show the back
            self.boxRow2Right_wndMain.add(
                self.btnDelete
            )
            self.boxRow4_wndMain.add(
                self.btnStrange,
                self.btnVague,
                self.btnFamiliar
            )
            
    def SaveWordListsAndCardsToFiles(self):
        global g_dictWordLists, g_dictCards
        jsonWordLists = json.dumps(g_dictWordLists)
        jsonCards = json.dumps(g_dictCards)
        with open(g_strWordlistsPath, 'w') as f:
            f.write(jsonWordLists)
        with open(g_strCardsPath, 'w') as f:
            f.write(jsonCards)

    def SaveConfiguration(self):
        (iX, iY) = self.main_window.position
        (iWidth, iHeight) = self.main_window.size
        json_str = json.dumps({
            "lang":         g_strDefaultLang,
            "cur_wordlist": g_strOpenedWordListName,
            "only_review":  g_bOnlyReview,
            "X":            iX,
            "Y":            iY,
            "width":        iWidth,
            "height":       iHeight
        })
        with open(g_strConfigPath, 'w') as f:
            f.write(json_str)
    
def main():
    return wordgets()