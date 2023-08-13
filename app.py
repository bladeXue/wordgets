import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW, LEFT, CENTER, RIGHT, HIDDEN, VISIBLE, NORMAL, BOLD, TRANSPARENT
import os
import platform
import sys
import json
import re
import copy
from functools import partial
import sqlite3
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import requests
from supermemo2 import SMTwo
from datetime import datetime
import random
from typing import Union

# Global variables
g_strRunPath        = ""
g_strLangPath       = ""
g_strDBPath         = ""
g_strConfigPath     = ""
g_strCardsPath      = ""
g_strWordlistsPath  = ""
g_dictCards         = {}
g_dictWordLists     = {}
g_EmptyHtml         = '<!DOCTYPE html> <html> <head> <title> </title> </head> <body> </body> </html>' #Cannot be "". Otherwise, it will not work
g_bIsMobile         = False

#Config
g_strDefaultLang        = "English"
g_strOpenedWordListName = ""
g_bOnlyReview           = False

# Color scheme
g_clrBg                     = "#EEF4F9"
g_clrPressedNavigationBtn   = "#E6EBF0"
g_clrStrange                = "#CFBAF0"
g_clrVague                  = "#FDE4CF"
g_clrFamiliar               = "#B9FBC0"
g_clrDelete                 = "#8EECF5"
g_clrOpenedTabText          = "#1651AA"


# Keywords for text display
g_strFront      = ""
g_strBack       = ""
g_strOnesided   = ""
g_strFrontend   = ""
g_strBackend    = ""
g_strFilePath   = ""
g_strCard1      = ""
g_strCard2      = ""

# Widgets texts
g_strLblNewWordsPerGroup                = ""
g_strLblOldWordsPerGroup                = ""
g_strLblReviewPolicy                    = ""
g_strCboReviewPolicy_item_LearnFirst    = ""
g_strCboReviewPolicy_item_Random        = ""
g_strCboReviewPolicy_item_ReviewFirst   = ""

# Messagebox
g_strErrorDialogTitle                       = ""
g_strErrorDialogMsg_EmptyName               = ""
g_strErrorDialogMsg_RepetitiveName          = ""
g_strErrorDialogMsg_InvalidFileName         = ""
g_strErrorDialogMsg_AssociatedWordList      = ""
g_strQuestionDialogTitle                    = ""
g_strQuestionDialogMsg_Delete               = ""
g_strErrorDialogMsg_InvalidFile             = ""
g_strErrorDialogMsg_NoWordList              = ""
g_strErrorDialogMsg_Card1IsNull             = ""
g_strErrorDialogMsg_Card1SameAsCard2        = ""
g_strQuestionDialogMsg_NotFoundWordList     = ""
g_strInfoDialogTitle                        = ""
g_strInfoDialogMsg_Complete                 = ""
g_strErrorDialogMsg_DownloadFailed          = ""
g_strInfoDialogMsg_MissionComplete          = ""
g_strInfoDialogMsg_ContinueToStudyNewWords  = ""
g_strQuestionDialogMsg_NoFurtherReview      = ""

#Temporary global
g_strTmpNameTxt             = ""
g_bTmpEditNewCard           = True
g_bTmpEditNewWordList       = True
g_iTmpPrevWordNo            = -1
g_iTmpPrevCardType          = -1
g_bSwitchContinueNewWords   = True
g_iTmpCntStudiedNewWords    = 0
g_iTmpCntStudiedOldWords    = 0 # It can be over than real total of old words.

# Functions
def GetOS() -> str:
    strOS = platform.system() # If Windows, return Windows
    if strOS == 'Linux':
        if str(hasattr(sys, 'getandroidapilevel')):
            strOS = 'Android'
        else:
            strOS = 'Linux'
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
        global g_strRunPath, g_strLangPath, g_strDBPath, g_strConfigPath, g_strCardsPath, g_strWordlistsPath
        g_strRunPath        = self.paths.app.absolute()
        g_strLangPath       = os.path.join(g_strRunPath, "resources/languages.json")
        g_strDBPath         = os.path.join(g_strRunPath, "resources/wordgets_stat.db")
        g_strConfigPath     = os.path.join(g_strRunPath, "resources/configuration.json")
        g_strCardsPath      = os.path.join(g_strRunPath, "resources/wordlists.json")
        g_strWordlistsPath  = os.path.join(g_strRunPath, "resources/cards.json")

        self.main_window    = toga.MainWindow(title=self.formal_name)
        boxMain             = toga.Box(style = Pack(direction = COLUMN, background_color = g_clrBg))
        boxNavigationBar    = toga.Box(style = Pack(direction = ROW, padding_top = 5, padding_bottom = 5, alignment = CENTER))
        
        self.btnIndex_boxNavigateBar       = toga.Button("",style = Pack(flex = 1), on_press = self.cbBtnIndexOnPress)
        self.btnSettings_boxNavigateBar    = toga.Button("",style = Pack(flex = 1), on_press = self.cbBtnSettingsOnPress)
        self.btnLibrary_boxNavigateBar     = toga.Button("",style = Pack(flex = 1), on_press = self.cbBtnLibraryOnPress)
        
        boxNavigationBar.add(
            self.btnIndex_boxNavigateBar,
            self.btnLibrary_boxNavigateBar,
            self.btnSettings_boxNavigateBar
        )
        self.boxBody             = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        
        boxMain.add(
            boxNavigationBar,
            self.boxBody
        )

        # Index body
        self.boxIndexBody = toga.Box(style=Pack(direction = COLUMN, flex = 1))
        boxIndexBody_Row1 = toga.Box(style = Pack(direction = ROW, alignment = CENTER))
        self.cboCurWordList_boxIndexBody = toga.Selection(style = Pack(flex = 1))
        self.chkReviewOnly_boxIndexBody = toga.Switch("", style= Pack(background_color = TRANSPARENT), on_change = self.cbChangeChkReviewOnly)
        boxIndexBody_Row1.add(
            self.cboCurWordList_boxIndexBody,
            self.chkReviewOnly_boxIndexBody
        )
        boxIndexBody_Row2 = toga.Box(style= Pack(direction = ROW))
        self.lblLearningProgressItem_boxIndexBody = toga.Label("", style = Pack(background_color = TRANSPARENT))
        self.lblLearningProgressValue_boxIndexBody = toga.Label("", style = Pack(background_color = TRANSPARENT))
        boxIndexBody_Row2.add(
            self.lblLearningProgressItem_boxIndexBody,
            self.lblLearningProgressValue_boxIndexBody
        )
        boxIndexBody_Row3 = toga.Box(style= Pack(direction = ROW, alignment = CENTER))
        boxIndexBody_Row3Hidden = toga.Box(style=Pack(direction = COLUMN))
        btnVoidForRemainRow3_boxIndexBody = toga.Button("",style=Pack(visibility = HIDDEN, width = 1))
        boxIndexBody_Row3Hidden.add(
            btnVoidForRemainRow3_boxIndexBody
        )
        self.boxIndexBody_Row3Blank = toga.Box(style=Pack(direction = COLUMN, flex = 1))
        boxIndexBody_Row3.add(
            boxIndexBody_Row3Hidden,
            self.boxIndexBody_Row3Blank
        )

        self.wbWord_boxIndexBody = toga.WebView(style = Pack(direction = COLUMN, flex = 1))

        boxIndexBody_Row5Border = toga.Box(style=Pack(direction = ROW))
        boxIndexBody_Row5Hidden = toga.Box(style=Pack(direction = ROW))
        btnVoidForRemainRow5_boxIndexBody = toga.Button("",style=Pack(visibility = HIDDEN, width = 1)) #If not, the box will not appear
        boxIndexBody_Row5Hidden.add(
            btnVoidForRemainRow5_boxIndexBody
        )
        self.boxIndexBody_Row5 = toga.Box(style=Pack(direction = ROW, flex = 1))
        boxIndexBody_Row5Border.add(
            boxIndexBody_Row5Hidden,
            self.boxIndexBody_Row5
        )

        self.boxIndexBody.add(
            boxIndexBody_Row1,
            boxIndexBody_Row2,
            boxIndexBody_Row3,
            self.wbWord_boxIndexBody,
            boxIndexBody_Row5Border
        )

        # Buttons
        self.btnStrange = toga.Button("", style = Pack(flex = 1, background_color = g_clrStrange))
        self.btnVague   = toga.Button("", style = Pack(flex = 1, background_color = g_clrVague))
        self.btnFamiliar= toga.Button("", style = Pack(flex = 1, background_color = g_clrFamiliar))
        self.btnDelete  = toga.Button("🗑️", style = Pack(flex = 1, background_color = g_clrDelete))
        self.btnShowBack= toga.Button("", style = Pack(flex = 1))

        self.btnShowBack.on_press = self.cbShowBack
        self.btnStrange.on_press = partial(self.cbNextCard, 0)
        self.btnVague.on_press = partial(self.cbNextCard, 2)
        self.btnFamiliar.on_press = partial(self.cbNextCard, 4)
        self.btnDelete.on_press = self.cbNoFurtherReviewThisWord

        # Settngs body
        self.boxSettingsBody = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        boxSettingsBody_Row1 = toga.Box(style = Pack(direction = ROW, alignment = CENTER))

        boxSettingsBody_Row1Col1 = toga.Box(style = Pack(direction = ROW, width = 150))
        self.lblLang_boxSettingsBody        = toga.Label("", style = Pack(background_color = TRANSPARENT))
        lblLangEmoji_boxSettingsBody   = toga.Label("🌐", style = Pack(background_color = TRANSPARENT))

        boxSettingsBody_Row1Col1.add(
            self.lblLang_boxSettingsBody,
            lblLangEmoji_boxSettingsBody
        )

        boxSettingsBody_Row1Col2 = toga.Box(style = Pack(direction = ROW))
        self.cboChangeLang_boxSettingsBody = toga.Selection(style = Pack(width = 300))
        boxSettingsBody_Row1Col2.add(
            self.cboChangeLang_boxSettingsBody
        )

        boxSettingsBody_Row1.add(
            boxSettingsBody_Row1Col1,
            boxSettingsBody_Row1Col2
        )
        
        self.boxSettingsBody.add(
            boxSettingsBody_Row1
        )
        
        # Library body
        self.boxLibraryBody = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        boxLibraryBodyRow1 = toga.Box(style = Pack(direction = ROW, flex = 1))
        boxLibraryBodyRow1Col1 = toga.Box(style = Pack(direction = COLUMN, width = 150))
        self.lblCards_boxLibraryBody = toga.Label("", style = Pack(background_color = TRANSPARENT))
        self.btnAddACard_boxLibraryBody = toga.Button("➕", style = Pack(flex = 1), on_press = self.cbAddACardOnPress)
        boxLibraryBodyRow1Col1.add(
            self.lblCards_boxLibraryBody,
            self.btnAddACard_boxLibraryBody
        )
        sbLibraryBodyRow1Col2 = toga.ScrollContainer(style = Pack(direction = COLUMN, flex = 1, background_color = TRANSPARENT))
        self.boxCards_boxLibraryBody = toga.Box(style = Pack(direction = COLUMN))
        sbLibraryBodyRow1Col2.content = self.boxCards_boxLibraryBody
        boxLibraryBodyRow1.add(
            boxLibraryBodyRow1Col1,
            sbLibraryBodyRow1Col2
        )
        
        boxLibraryBodyRow2 = toga.Box(style = Pack(direction = ROW, flex = 1))
        boxLibraryBodyRow2Col1 = toga.Box(style = Pack(direction = COLUMN, width = 150))
        self.lblWordLists_boxLibraryBody = toga.Label("", style = Pack(background_color = TRANSPARENT))
        self.btnAddAWordList_boxLibraryBody = toga.Button("➕", style = Pack(flex = 1), on_press = self.cbAddAWordListdOnPress)
        boxLibraryBodyRow2Col1.add(
            self.lblWordLists_boxLibraryBody,
            self.btnAddAWordList_boxLibraryBody
        )
        sbLibraryBodyRow2Col2 = toga.ScrollContainer(style = Pack(direction = COLUMN, flex = 1, background_color = TRANSPARENT))
        self.boxWordLists_boxLibraryBody = toga.Box(style = Pack(direction = COLUMN))
        sbLibraryBodyRow2Col2.content = self.boxWordLists_boxLibraryBody
        boxLibraryBodyRow2.add(
            boxLibraryBodyRow2Col1,
            sbLibraryBodyRow2Col2
        )
        self.btnApplyChanges_boxLibraryBody = toga.Button("", on_press = self.cbApplyChangesOnPress)
        self.boxLibraryBody.add(
            boxLibraryBodyRow1,
            boxLibraryBodyRow2,
            self.btnApplyChanges_boxLibraryBody
        )

        # Body of editing a card   
        self.boxEditCardBody = toga.Box(style = Pack(direction = COLUMN, flex = 1))

        sbEditCardBodyRow1 = toga.ScrollContainer(style = Pack(direction = COLUMN, flex = 1, background_color = TRANSPARENT), horizontal = False)

        self.lblEditCard_boxEditCardBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1, font_weight = BOLD))
        boxEditCardBodyRow1Row2 = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        self.lblCardName_boxEditCardBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1))
        self.txtCardName_boxEditCardBody = toga.TextInput(style = Pack(flex = 1), on_change = self.cbValidateNameTxt)
        boxEditCardBodyRow1Row2.add(
            self.lblCardName_boxEditCardBody,
            self.txtCardName_boxEditCardBody
        )
        boxEditCardBodyRow1Row3 = toga.Box(style = Pack(direction = COLUMN,flex = 1))
        self.lblCardFrontFrontend_boxEditCardBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1))
        self.txtCardFrontFrontend_boxEditCardBody = toga.TextInput(style = Pack(flex = 1))
        self.lblCardFrontBackend_boxEditCardBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1))
        self.txtCardFrontBackend_boxEditCardBody = toga.TextInput(style = Pack(flex = 1))
        boxEditCardBodyRow1Row3.add(
            self.lblCardFrontFrontend_boxEditCardBody,
            self.txtCardFrontFrontend_boxEditCardBody,
            self.lblCardFrontBackend_boxEditCardBody,
            self.txtCardFrontBackend_boxEditCardBody
        )
        self.chkEnableCardBack = toga.Switch("", style = Pack(background_color = TRANSPARENT, flex = 1), value = False, on_change = self.cbEnableCardBackOnChange)
        self.boxEditCardBodyRow1Row5 = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        self.lblCardBackFrontend_boxEditCardBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1))
        self.txtCardBackFrontend_boxEditCardBody = toga.TextInput(style = Pack(flex = 1))
        self.lblCardBackBackend_boxEditCardBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1))
        self.txtCardBackBackend_boxEditCardBody = toga.TextInput(style = Pack(flex = 1))
        self.boxEditCardBodyRow1Row5.add(
            self.lblCardBackFrontend_boxEditCardBody,
            self.txtCardBackFrontend_boxEditCardBody,
            self.lblCardBackBackend_boxEditCardBody,
            self.txtCardBackBackend_boxEditCardBody
        )

        sbEditCardBodyRow1.content = toga.Box(
            children=[
                self.lblEditCard_boxEditCardBody,
                boxEditCardBodyRow1Row2,
                boxEditCardBodyRow1Row3,
                self.chkEnableCardBack,
                self.boxEditCardBodyRow1Row5
            ],style = Pack(direction = COLUMN)
        )

        self.btnSaveCard_boxEditCardBody = toga.Button("", on_press = self.cbSaveCardOnPress)
        self.boxEditCardBody.add(
            sbEditCardBodyRow1,
            self.btnSaveCard_boxEditCardBody
        )
        self.boxEditCardBodyRow1Row5.style.update(visibility = HIDDEN)

        # Body of editing a word list
        self.boxEditWordListBody = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        sbEditWordListBodyRow1 = toga.ScrollContainer(style = Pack(direction = COLUMN, flex = 1, background_color = TRANSPARENT), horizontal = False)
        self.lblEditWordlist_boxEditWordListBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1, font_weight = BOLD))
        boxEditWordListBodyRow1Row2 = toga.Box(style = Pack(direction = COLUMN, flex = 1))
        self.lblWordListName_boxEditWordListBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1))
        self.txtWordListName_boxEditWordListBody = toga.TextInput(style = Pack(flex = 1), on_change = self.cbValidateNameTxt)
        boxEditWordListBodyRow1Row2.add(
            self.lblWordListName_boxEditWordListBody,
            self.txtWordListName_boxEditWordListBody
        )
        boxEditWordListBodyRow1Row3 = toga.Box(style = Pack(direction = COLUMN,flex = 1))
        self.lblWordListXlsx_boxEditWordListBody = toga.Label("", style = Pack(background_color = TRANSPARENT, flex = 1))
        self.txtWordListXlsx_boxEditWordListBody = toga.TextInput(style = Pack(flex = 1))
        
        boxEditWordListBodyRow1Row3.add(
            self.lblWordListXlsx_boxEditWordListBody,
            self.txtWordListXlsx_boxEditWordListBody,
        )

        sbEditWordListBodyRow1.content = toga.Box(
            children=[
                self.lblEditWordlist_boxEditWordListBody,
                boxEditWordListBodyRow1Row2,
                boxEditWordListBodyRow1Row3
            ],style = Pack(direction = COLUMN)
        )

        self.btnSaveWordList_boxEditWordListBody = toga.Button("", on_press = self.cbSaveWordListOnPress)
        self.boxEditWordListBody.add(
            sbEditWordListBodyRow1,
            self.btnSaveWordList_boxEditWordListBody
        )

        # Read language file
        self.m_dictLanguages = {}
        with open(g_strLangPath, "r") as f:
            self.m_dictLanguages = json.load(f)

        global g_bIsMobile, g_strDefaultLang, g_strOpenedWordListName, g_bOnlyReview
        if GetOS() in ['Windows','Linux','MacOS']:
            g_bIsMobile = False

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
                    #Current lack resolution getter
                    if 0 <= t_iX <= t_iX + t_iWidth  and\
                        0 <= t_iY <= t_iY + t_iHeight :
                        iX = t_iX
                        iY = t_iY
                        iWidth = t_iWidth
                        iHeight = t_iHeight
        
        self.main_window.content = boxMain
        self.main_window.position = (iX, iY)
        self.main_window.size = (iWidth, iHeight)
        
        #Initialize
        self.cbBtnIndexOnPress(None)
        self.main_window.show()

        self.cboChangeLang_boxSettingsBody.items = list(self.m_dictLanguages.keys())
        self.cboChangeLang_boxSettingsBody.value = g_strDefaultLang

        self.cboChangeLang_boxSettingsBody.on_select = self.cbChangeLangOnSelect
        self.ChangeLangAccordingToDefaultLang()

        self.chkReviewOnly_boxIndexBody.value = g_bOnlyReview
        
        self.main_window.info_dialog("test",g_strWordlistsPath)
        if os.path.exists(g_strWordlistsPath) == True:
            with open(g_strWordlistsPath, "r") as f:
                global g_dictWordLists
                g_dictWordLists = json.load(f)
                self.main_window.info_dialog("test",str(g_dictWordLists))
        self.main_window.info_dialog("test",g_strCardsPath)
        if os.path.exists(g_strCardsPath) == True:
            with open(g_strCardsPath, "r") as f:
                global g_dictCards
                g_dictCards = json.load(f)
                self.main_window.info_dialog("test",str(g_dictCards))
        
        self.wbWord_boxIndexBody.set_content("NO_CONTENT", g_EmptyHtml)

        self.m_lsRememberSeq = []

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
            
            self.cboCurWordList_boxIndexBody.items = lsValidWordlists
            self.cboCurWordList_boxIndexBody.value = g_strOpenedWordListName


            self.cboCurWordList_boxIndexBody.on_select = self.cbChangeCurWordList
            
            self.GenerateNewWordsSeqOfCurrentWordList()
            self.cbNextCard(-1, None)
        

    # Callback functions
    def cbBtnIndexOnPress(self, widget):
        while len(self.boxBody.children) != 0:
            self.boxBody.remove(self.boxBody.children[0])
        self.btnIndex_boxNavigateBar.style.update(font_weight = BOLD, background_color = g_clrPressedNavigationBtn, color = g_clrOpenedTabText)
        self.btnSettings_boxNavigateBar.style.update(font_weight = NORMAL, background_color = g_clrBg, color = 'black')
        self.btnLibrary_boxNavigateBar.style.update(font_weight = NORMAL, background_color = g_clrBg, color = 'black')
        self.boxBody.add(
            self.boxIndexBody
        )
        self.SaveConfiguration()

    def cbBtnSettingsOnPress(self, widget):
        while len(self.boxBody.children) != 0:
            self.boxBody.remove(self.boxBody.children[0])
        self.btnIndex_boxNavigateBar.style.update(font_weight = NORMAL, background_color = g_clrBg, color = 'black')
        self.btnSettings_boxNavigateBar.style.update(font_weight = BOLD, background_color = g_clrPressedNavigationBtn, color = g_clrOpenedTabText)
        self.btnLibrary_boxNavigateBar.style.update(font_weight = NORMAL, background_color = g_clrBg, color = 'black')
        self.boxBody.add(
            self.boxSettingsBody
        )
        self.SaveConfiguration()

    def cbBtnLibraryOnPress(self, widget):
        while len(self.boxBody.children) != 0:
            self.boxBody.remove(self.boxBody.children[0])
        self.btnIndex_boxNavigateBar.style.update(font_weight = NORMAL, background_color = g_clrBg, color = 'black')
        self.btnSettings_boxNavigateBar.style.update(font_weight = NORMAL, background_color = g_clrBg, color = 'black')
        self.btnLibrary_boxNavigateBar.style.update(font_weight = BOLD, background_color = g_clrPressedNavigationBtn, color = g_clrOpenedTabText)
        self.boxBody.add(
            self.boxLibraryBody
        )
        self.SaveConfiguration()
        
        self.m_dictTmpCards = copy.deepcopy(g_dictCards)
        self.m_dictTmpWordLists = copy.deepcopy(g_dictWordLists)
        self.FreshCards()
        self.FreshWordLists()


    def cbChangeLangOnSelect(self, widget):
        global g_strDefaultLang
        g_strDefaultLang = widget.value
        self.ChangeLangAccordingToDefaultLang()
        
    def cbValidateNameTxt(self, widget):
        global g_strTmpNameTxt
        if widget.value == "":
            g_strTmpNameTxt = ""
            return
        if not re.match(r"^[a-zA-Z_][a-zA-Z0-9_]*$", widget.value):
            widget.value = g_strTmpNameTxt
        if len(widget.value) > 10:
            widget.value = g_strTmpNameTxt
        g_strTmpNameTxt = widget.value
        
    
    def cbAddACardOnPress(self, widget):
        while len(self.boxBody.children) != 0:
            self.boxBody.remove(self.boxBody.children[0])
        self.boxBody.add(self.boxEditCardBody)
        while len(self.boxBody.children) != 0: # A BUG: It must be execute twice at the first time, or the scrollcontainer opened at the first time could be incomplete
            self.boxBody.remove(self.boxBody.children[0])
        self.boxBody.add(self.boxEditCardBody)
        self.txtCardName_boxEditCardBody.value = ""
        self.txtCardFrontFrontend_boxEditCardBody.value = ""
        self.txtCardFrontBackend_boxEditCardBody.value = ""
        self.txtCardBackFrontend_boxEditCardBody.value = ""
        self.txtCardBackBackend_boxEditCardBody.value = ""
        global g_bTmpEditNewCard
        g_bTmpEditNewCard = True
        self.txtCardName_boxEditCardBody.readonly = False
    
    def cbAddAWordListdOnPress(self, widget):
        while len(self.boxBody.children) != 0:
            self.boxBody.remove(self.boxBody.children[0])
        self.boxBody.add(self.boxEditWordListBody)
        while len(self.boxBody.children) != 0: # A BUG: It must be execute twice at the first time, or the scrollcontainer opened at the first time could be incomplete
            self.boxBody.remove(self.boxBody.children[0])
        self.boxBody.add(self.boxEditWordListBody)
        
        self.txtWordListName_boxEditWordListBody.value = ""
        self.txtWordListXlsx_boxEditWordListBody.value = ""
        global g_bTmpEditNewWordList
        g_bTmpEditNewWordList = True
        self.txtWordListName_boxEditWordListBody.readonly = False

    def cbSaveCardOnPress(self, widget):
        strCardCname = self.txtCardName_boxEditCardBody.value
        if strCardCname == "":
            self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_EmptyName)
            return
        if strCardCname in self.m_dictTmpCards and g_bTmpEditNewCard == True:
            self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_RepetitiveName)
            return
        
        strFrontFrontendFilePath = self.txtCardFrontFrontend_boxEditCardBody.value
        strFrontBackendFilePath = self.txtCardFrontBackend_boxEditCardBody.value
        strBackFrontendFilePath = self.txtCardBackFrontend_boxEditCardBody.value
        strBackBackendFilePath = self.txtCardBackBackend_boxEditCardBody.value

        bEnableBack = self.chkEnableCardBack.value
        if bEnableBack == False:
            if strFrontFrontendFilePath == "" or strFrontBackendFilePath == "":
                self.main_window.error_dialog(title = g_strErrorDialogTitle, message = g_strErrorDialogMsg_InvalidFileName)
                return
            else:
                if strFrontFrontendFilePath[:4] == "http":
                    try:
                        response = requests.get(strFrontFrontendFilePath, stream = True)
                        filename = strCardCname+"-front.html"
                        strFileWriteTo = os.path.join(g_strRunPath, "download/"+filename)
                        with open(strFileWriteTo, "wb") as f:
                            for chunk in response.iter_content(chunk_size = 1024*1024):
                                if chunk:
                                    f.write(chunk)
                        strFrontFrontendFilePath = strFileWriteTo
                    except Exception:
                        self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_DownloadFailed)
                        return
                    
                if strFrontBackendFilePath[:4] == "http":
                    try:
                        response = requests.get(strFrontBackendFilePath, stream = True)
                        filename = strCardCname+"-front.py"
                        strFileWriteTo = os.path.join(g_strRunPath, "download/"+filename)
                        with open(strFileWriteTo, "wb") as f:
                            for chunk in response.iter_content(chunk_size = 1024*1024):
                                if chunk:
                                    f.write(chunk)
                        strFrontBackendFilePath = strFileWriteTo
                    except Exception:
                        self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_DownloadFailed)
                        return
                    
                self.m_dictTmpCards[strCardCname] = {
                    "single-sided": {
                        "frontend": strFrontFrontendFilePath,
                        "backend": strFrontBackendFilePath
                    }
                }
        else:
            if strFrontFrontendFilePath == "" or strFrontBackendFilePath == "" or strBackFrontendFilePath == "" or strBackBackendFilePath == "":
                self.main_window.error_dialog(title = g_strErrorDialogTitle, message = g_strErrorDialogMsg_InvalidFileName)
                return
            else:
                if strFrontFrontendFilePath[:4] == "http":
                    try:
                        response = requests.get(strFrontFrontendFilePath, stream = True)
                        filename = strCardCname+"-front.html"
                        strFileWriteTo = os.path.join(g_strRunPath, "download/"+filename)
                        with open(strFileWriteTo, "wb") as f:
                            for chunk in response.iter_content(chunk_size = 1024*1024):
                                if chunk:
                                    f.write(chunk)
                        strFrontFrontendFilePath = strFileWriteTo
                    except Exception:
                        self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_DownloadFailed)
                        return
                    
                if strFrontBackendFilePath[:4] == "http":
                    try:
                        response = requests.get(strFrontBackendFilePath, stream = True)
                        filename = strCardCname+"-front.py"
                        strFileWriteTo = os.path.join(g_strRunPath, "download/"+filename)
                        with open(strFileWriteTo, "wb") as f:
                            for chunk in response.iter_content(chunk_size = 1024*1024):
                                if chunk:
                                    f.write(chunk)
                        strFrontBackendFilePath = strFileWriteTo
                    except Exception:
                        self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_DownloadFailed)
                        return
                if strBackFrontendFilePath[:4] == "http":
                    try:
                        response = requests.get(strBackFrontendFilePath, stream = True)
                        filename = strCardCname+"-back.html"
                        strFileWriteTo = os.path.join(g_strRunPath, "download/"+filename)
                        with open(strFileWriteTo, "wb") as f:
                            for chunk in response.iter_content(chunk_size = 1024*1024):
                                if chunk:
                                    f.write(chunk)
                        strBackFrontendFilePath = strFileWriteTo
                    except Exception:
                        self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_DownloadFailed)
                        return
                    
                if strBackBackendFilePath[:4] == "http":
                    try:
                        response = requests.get(strBackBackendFilePath, stream = True)
                        filename = strCardCname+"-back.py"
                        strFileWriteTo = os.path.join(g_strRunPath, "download/"+filename)
                        with open(strFileWriteTo, "wb") as f:
                            for chunk in response.iter_content(chunk_size = 1024*1024):
                                if chunk:
                                    f.write(chunk)
                        strBackBackendFilePath = strFileWriteTo
                    except Exception:
                        self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_DownloadFailed)
                        return

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
        self.BackToLibraryAndFresh()

    def cbSaveWordListOnPress(self, widget):
        strWordListCname = self.txtWordListName_boxEditWordListBody.value
        if strWordListCname == "":
            self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_EmptyName)
            return
        if strWordListCname in self.m_dictTmpWordLists and g_bTmpEditNewWordList == True:
            self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_RepetitiveName)
            return 
        strExcelFilePath = self.txtWordListXlsx_boxEditWordListBody.value

        if strExcelFilePath == "":
            self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_InvalidFileName)
            return
        else:
            if strExcelFilePath[:4] == "http":
                    try:
                        response = requests.get(strExcelFilePath, stream = True)
                        filename = strWordListCname+".xlsx"
                        strFileWriteTo = os.path.join(g_strRunPath, "download/"+filename)
                        with open(strFileWriteTo, "wb") as f:
                            for chunk in response.iter_content(chunk_size = 1024*1024):
                                if chunk:
                                    f.write(chunk)
                        strExcelFilePath = strFileWriteTo
                    except Exception:
                        self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_DownloadFailed)
                        return
            self.m_dictTmpWordLists[strWordListCname] = {
                "FilePath": strExcelFilePath,
                "CardType1": "",
                "CardType2": "",
                "NewWordsPerGroup": 20,
                "OldWordsPerGroup": 20,
                "policy": 0
            }
        
        self.BackToLibraryAndFresh()
        
    def cbEnableCardBackOnChange(self, widget):
        if widget.value == True:
            self.boxEditCardBodyRow1Row5.style.update(visibility = VISIBLE)
        else:
            self.boxEditCardBodyRow1Row5.style.update(visibility = HIDDEN)

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
            self.main_window.error_dialog(title = g_strErrorDialogTitle, message = g_strErrorDialogMsg_AssociatedWordList)
            return
        self.m_dictTmpCards.pop(strCardName)
        self.FreshCards()
        self.FreshWordLists()

    def cbEditExistedCard(self, strCardName, widget):
        self.txtCardName_boxEditCardBody.readonly = True
        self.txtCardName_boxEditCardBody.value = strCardName
        if len(self.m_dictTmpCards[strCardName]) == 1:
            self.txtCardFrontFrontend_boxEditCardBody.value = self.m_dictTmpCards[strCardName]['single-sided']['frontend']
            self.txtCardFrontBackend_boxEditCardBody.value = self.m_dictTmpCards[strCardName]['single-sided']['backend']
            self.chkEnableCardBack.value = False
            self.txtCardBackFrontend_boxEditCardBody.value = ""
            self.txtCardBackBackend_boxEditCardBody.value = ""
        else:
            self.txtCardFrontFrontend_boxEditCardBody.value = self.m_dictTmpCards[strCardName]['front']['frontend']
            self.txtCardFrontBackend_boxEditCardBody.value = self.m_dictTmpCards[strCardName]['front']['backend']
            self.chkEnableCardBack.value = True
            self.txtCardBackFrontend_boxEditCardBody.value = self.m_dictTmpCards[strCardName]['back']['frontend']
            self.txtCardBackBackend_boxEditCardBody.value = self.m_dictTmpCards[strCardName]['back']['backend']
        global g_bTmpEditNewCard
        g_bTmpEditNewCard = False
        while len(self.boxBody.children) != 0:
            self.boxBody.remove(self.boxBody.children[0])
        self.boxBody.add(self.boxEditCardBody)

    async def cbDeleteWordList(self, widget):
        if await self.main_window.question_dialog(title = g_strQuestionDialogTitle, message = g_strQuestionDialogMsg_Delete):
            self.m_dictTmpWordLists.pop(widget.id)  # An unreasonable solution
            self.FreshCards()
            self.FreshWordLists()
        else:
            return
    
    def cbChangeReviewPolicy(self, strWordList, widget):
        self.m_dictTmpWordLists[strWordList]['policy'] = widget.items.index(widget.value)

    def cbChangeNewWordsPerGroup(self, strWordList, widget):
        try:
            self.m_dictTmpWordLists[strWordList]['NewWordsPerGroup'] = int(widget.value)
        except Exception:
            self.m_dictTmpWordLists[strWordList]['NewWordsPerGroup'] = 1

    def cbChangeOldWordsPerGroup(self, strWordList, widget):
        try:
            self.m_dictTmpWordLists[strWordList]['OldWordsPerGroup'] = int(widget.value)
        except Exception:
            self.m_dictTmpWordLists[strWordList]['OldWordsPerGroup'] = 1
    
    def cbChangeWordListCardType(self, strWordList, iDeleteCardType, widget):
        if iDeleteCardType == 1:
            if widget.value == None:
                self.m_dictTmpWordLists[strWordList]['CardType1'] = ""
            else:
                self.m_dictTmpWordLists[strWordList]['CardType1'] = widget.value
        elif iDeleteCardType == 2:
            if widget.value == None:
                self.m_dictTmpWordLists[strWordList]['CardType2'] = ""
            else:
                self.m_dictTmpWordLists[strWordList]['CardType2'] = widget.value

    def cbEditExistedWordList(self, strWordListName, widget):
        self.txtWordListName_boxEditWordListBody.readonly = True
        self.txtWordListName_boxEditWordListBody.value = strWordListName
        self.txtWordListXlsx_boxEditWordListBody.value = self.m_dictTmpWordLists[strWordListName]['FilePath']
        global g_bTmpEditNewWordList
        g_bTmpEditNewWordList = False
        while len(self.boxBody.children) != 0:
            self.boxBody.remove(self.boxBody.children[0])
        self.boxBody.add(self.boxEditWordListBody)
    
    async def cbApplyChangesOnPress(self, widget):
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
            self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_InvalidFile)
            return
        
        # Check if there is no word list
        if len(self.m_dictTmpWordLists) == 0:
            self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_NoWordList)
            return
        
        # Check all word list is available
        for eachWordList in self.m_dictTmpWordLists:
            if self.m_dictTmpWordLists[eachWordList]['CardType1'] == "":
                self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_Card1IsNull)
                return
            else:
                if self.m_dictTmpWordLists[eachWordList]['CardType2'] == self.m_dictTmpWordLists[eachWordList]['CardType1']:
                    self.main_window.error_dialog(g_strErrorDialogTitle, g_strErrorDialogMsg_Card1SameAsCard2)
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
                if await self.main_window.question_dialog(g_strQuestionDialogTitle, message = g_strQuestionDialogMsg_NotFoundWordList % eachWordListInDB):
                    cursor.execute(f"DELETE FROM statistics WHERE wordlist='{eachWordListInDB}'")  
                    lsAllWordListsInDBRemained.remove(eachWordListInDB)
                else:
                    continue
        
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

        self.main_window.info_dialog(title = g_strInfoDialogTitle, message = g_strInfoDialogMsg_Complete)

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

        self.cboCurWordList_boxIndexBody.items = lsValidWordlists
        self.cboCurWordList_boxIndexBody.value = g_strOpenedWordListName

        self.GenerateNewWordsSeqOfCurrentWordList()
        self.cbNextCard(-1, None)

        self.cbBtnIndexOnPress(None)

    def cbNextCard(self, iQuality, widget): # iQuality = -1 if there is no need to update record
        self.ChangeWordLearningBtns(0)
        conn = sqlite3.connect(g_strDBPath)
        cursor = conn.cursor()

        global g_bSwitchContinueNewWords, g_iTmpCntStudiedNewWords, g_iTmpCntStudiedOldWords, g_iTmpPrevWordNo, g_iTmpPrevCardType

        today = datetime.today().date()
        strToday = datetime.now().strftime('%Y-%m-%d')

        # Write old word record
        if iQuality != -1:
            if len(self.m_lsRememberSeq) != 0:
                if self.m_lsRememberSeq[0][0]== g_iTmpPrevWordNo and self.m_lsRememberSeq[0][1] == g_iTmpPrevCardType:
                    self.m_lsRememberSeq.pop(0)
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
        cursor.execute("SELECT COUNT(*) FROM statistics WHERE wordlist = ? AND review_date IS NOT NULL",(g_strOpenedWordListName,))
        iCntStudiedCards = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM statistics WHERE wordlist = ? AND review_date IS NULL",(g_strOpenedWordListName,))
        iCntRestNewCards = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM statistics WHERE wordlist = ? AND review_date <= ?", (g_strOpenedWordListName, strToday))
        iCntDueTimeCards = cursor.fetchone()[0]

        self.lblLearningProgressValue_boxIndexBody.text = str(iCntStudiedCards)+"("+ str(g_iTmpCntStudiedNewWords) + "," + str(g_iTmpCntStudiedOldWords) \
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
                self.wbWord_boxIndexBody.set_content("NO_CONTENT", g_EmptyHtml)
                conn.commit()
                conn.close()
                # If only review is true and rest new cards is not 0,
                if g_bOnlyReview == True and iCntRestNewCards > 0:
                    if g_iTmpCntStudiedNewWords == 0 and g_iTmpCntStudiedOldWords == 0: # A bug: if the program just runs, the info dialog will raise RuntimeError.
                        return
                    self.main_window.info_dialog(g_strInfoDialogTitle, g_strInfoDialogMsg_ContinueToStudyNewWords)
                    return
                # All old words have been studied.
                if g_iTmpCntStudiedNewWords == 0 and g_iTmpCntStudiedOldWords == 0: # A bug: if the program just runs, the info dialog will raise RuntimeError.
                    return
                self.main_window.info_dialog(g_strInfoDialogTitle, g_strInfoDialogMsg_MissionComplete)
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
                
                self.wbWord_boxIndexBody.set_content("NO_CONTENT", g_EmptyHtml)
                conn.commit()
                conn.close()
                if g_iTmpCntStudiedNewWords == 0 and g_iTmpCntStudiedOldWords == 0: # A bug: if the program just runs, the info dialog will raise RuntimeError.
                    return
                self.main_window.info_dialog(g_strInfoDialogTitle, g_strInfoDialogMsg_MissionComplete)
                return
            else:
                g_iTmpPrevWordNo = self.m_lsRememberSeq[0][0]
                g_iTmpPrevCardType = self.m_lsRememberSeq[0][1]
                g_iTmpCntStudiedNewWords += 1
            
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
            self.wbWord_boxIndexBody.set_content("wordgets", strRenderedHTML)

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
            self.wbWord_boxIndexBody.set_content("wordgets", strRenderedHTML)

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

        self.wbWord_boxIndexBody.set_content("wordgets", strRenderedHTML)

        self.ChangeWordLearningBtns(2)

    def cbChangeCurWordList(self, widget):
        global g_iTmpPrevWordNo, g_iTmpPrevCardType, g_iTmpCntStudiedNewWords, g_iTmpCntStudiedOldWords, g_strOpenedWordListName
        if widget.value == g_strOpenedWordListName:
            return
        if widget.value == None:
            g_strOpenedWordListName = ""
        else:
            g_strOpenedWordListName = widget.value
        g_iTmpPrevWordNo = -1
        g_iTmpPrevCardType = -1
        g_iTmpCntStudiedNewWords = 0
        g_iTmpCntStudiedOldWords = 0
        self.lblLearningProgressValue_boxIndexBody.text = ""
        self.GenerateNewWordsSeqOfCurrentWordList()
        self.cbNextCard(-1, None)
        self.SaveConfiguration()

    async def cbNoFurtherReviewThisWord(self, widget):
        if await self.main_window.question_dialog(title = g_strQuestionDialogTitle, message = g_strQuestionDialogMsg_NoFurtherReview):
            self.cbNextCard(5, None)
        else:
            return
    
    def cbChangeChkReviewOnly(self, widget):
        global g_bOnlyReview
        g_bOnlyReview = widget.value
        if g_bOnlyReview == False:
            self.cbNextCard(-1, None)

    # Functions
    def ChangeLangAccordingToDefaultLang(self):
        self.btnIndex_boxNavigateBar.text                   = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_INDEX_TEXT']
        self.btnSettings_boxNavigateBar.text                = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_SETTINGS_TEXT']
        self.btnLibrary_boxNavigateBar.text                 = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_LIBRARY_TEXT']
        self.lblLang_boxSettingsBody.text                   = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_LANG_TEXT']
        self.lblCards_boxLibraryBody.text                   = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_CARDS_TEXT']
        self.lblWordLists_boxLibraryBody.text               = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_WORDLISTS_TEXT']
        self.btnApplyChanges_boxLibraryBody.text            = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_APPLYCHANGES_TEXT']
        self.lblEditCard_boxEditCardBody.text               = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_EDITCARD_TEXT']
        self.lblCardName_boxEditCardBody.text               = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_CARDNAME_TEXT']
        self.txtCardName_boxEditCardBody.placeholder        = self.m_dictLanguages[g_strDefaultLang]['STR_TXT_PLACEHOLDER_LEGALNAME']
        self.lblCardFrontFrontend_boxEditCardBody.text      = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_CARDFRONT_FRONTEND_TEXT']
        self.lblCardFrontBackend_boxEditCardBody.text       = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_CARDFRONT_BACKEND_TEXT']
        self.lblCardBackFrontend_boxEditCardBody.text       = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_CARDBACK_FRONTEND_TEXT']
        self.lblCardBackBackend_boxEditCardBody.text        = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_CARDBACK_BACKEND_TEXT']
        self.chkEnableCardBack.text                         = self.m_dictLanguages[g_strDefaultLang]['STR_CHK_ENABLECARDBACK_TEXT']
        self.btnSaveCard_boxEditCardBody.text               = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_SAVECARD_TEXT']
        self.lblEditWordlist_boxEditWordListBody.text       = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_EDITWORDLIST_TEXT']
        self.lblWordListName_boxEditWordListBody.text       = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_WORDLISTNAME_TEXT']
        self.txtWordListName_boxEditWordListBody.placeholder= self.m_dictLanguages[g_strDefaultLang]['STR_TXT_PLACEHOLDER_LEGALNAME']
        self.lblWordListXlsx_boxEditWordListBody.text       = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_WORDLISTXLSX_TEXT']
        self.btnSaveWordList_boxEditWordListBody.text       = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_SAVEWORDLIST_TEXT']
        self.lblLearningProgressItem_boxIndexBody.text      = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_LEARNINGPROGRESSITEM_TEXT']
        self.chkReviewOnly_boxIndexBody.text                = self.m_dictLanguages[g_strDefaultLang]['STR_CHK_REVIEWONLY_TEXT']
        self.btnShowBack.text                               = self.m_dictLanguages[g_strDefaultLang]['STR_BTN_SHOWBACK_TEXT']
        global g_strFront, g_strBack, g_strOnesided, g_strFrontend, g_strBackend, g_strFilePath, g_strCard1, g_strCard2
        g_strFront                                          = self.m_dictLanguages[g_strDefaultLang]['STR_FRONT']
        g_strBack                                           = self.m_dictLanguages[g_strDefaultLang]['STR_BACK']
        g_strOnesided                                       = self.m_dictLanguages[g_strDefaultLang]['STR_ONESIDED']
        g_strFrontend                                       = self.m_dictLanguages[g_strDefaultLang]['STR_FRONTEND']
        g_strBackend                                        = self.m_dictLanguages[g_strDefaultLang]['STR_BACKEND']
        g_strFilePath                                       = self.m_dictLanguages[g_strDefaultLang]['STR_FILEPATH']
        g_strCard1                                          = self.m_dictLanguages[g_strDefaultLang]['STR_CARD_1']
        g_strCard2                                          = self.m_dictLanguages[g_strDefaultLang]['STR_CARD_2']
        global g_strErrorDialogTitle, g_strErrorDialogMsg_EmptyName, g_strErrorDialogMsg_InvalidFileName, g_strErrorDialogMsg_RepetitiveName,\
            g_strErrorDialogMsg_AssociatedWordList, g_strQuestionDialogTitle, g_strQuestionDialogMsg_Delete, g_strErrorDialogMsg_InvalidFile,\
            g_strErrorDialogMsg_NoWordList, g_strErrorDialogMsg_Card1IsNull, g_strErrorDialogMsg_Card1SameAsCard2,\
            g_strQuestionDialogMsg_NotFoundWordList, g_strInfoDialogTitle, g_strInfoDialogMsg_Complete, g_strErrorDialogMsg_DownloadFailed,\
            g_strInfoDialogMsg_MissionComplete, g_strInfoDialogMsg_ContinueToStudyNewWords, g_strQuestionDialogMsg_NoFurtherReview
        g_strErrorDialogTitle                               = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_TITLE']
        g_strErrorDialogMsg_EmptyName                       = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_EMPTYNAME']
        g_strErrorDialogMsg_RepetitiveName                  = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_REPETITIVENAME']
        g_strErrorDialogMsg_InvalidFileName                 = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_INVALIDFILENAME']
        g_strErrorDialogMsg_AssociatedWordList              = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_ASSOCIATEDWORDLIST']
        g_strQuestionDialogTitle                            = self.m_dictLanguages[g_strDefaultLang]['STR_QUESTIONDIALOG_TITLE']
        g_strQuestionDialogMsg_Delete                       = self.m_dictLanguages[g_strDefaultLang]['STR_QUESTIONDIALOG_MSG_DELETE']
        g_strErrorDialogMsg_InvalidFile                     = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_INVALIDFILE']
        g_strErrorDialogMsg_NoWordList                      = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_NOWORDLIST']
        g_strErrorDialogMsg_Card1IsNull                     = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_CARD1ISNULL']
        g_strErrorDialogMsg_Card1SameAsCard2                = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_CARD1SAMEASCARD2']
        g_strQuestionDialogMsg_NotFoundWordList             = self.m_dictLanguages[g_strDefaultLang]['STR_QUESTIONDIALOG_MSG_NOTFOUNDWORDLIST']
        g_strInfoDialogTitle                                = self.m_dictLanguages[g_strDefaultLang]['STR_INFODIALOG_TITLE']
        g_strInfoDialogMsg_Complete                         = self.m_dictLanguages[g_strDefaultLang]['STR_INFODIALOG_MSG_COMPLETE']
        g_strErrorDialogMsg_DownloadFailed                  = self.m_dictLanguages[g_strDefaultLang]['STR_ERRORDIALOG_MSG_DOWNLOADFAILED']
        g_strInfoDialogMsg_MissionComplete                  = self.m_dictLanguages[g_strDefaultLang]['STR_INFODIALOG_MSG_MISSIONCOMPLETE']
        g_strInfoDialogMsg_ContinueToStudyNewWords          = self.m_dictLanguages[g_strDefaultLang]['STR_INFODIALOG_MSG_CONTINUETOSTUDYINGNEWWORDS']
        g_strQuestionDialogMsg_NoFurtherReview              = self.m_dictLanguages[g_strDefaultLang]['STR_QUESTIONDIALOG_MSG_NOFURTHERREVIEW']
        global g_strLblNewWordsPerGroup, g_strLblOldWordsPerGroup, g_strLblReviewPolicy, g_strCboReviewPolicy_item_LearnFirst, \
            g_strCboReviewPolicy_item_Random, g_strCboReviewPolicy_item_ReviewFirst
        g_strLblNewWordsPerGroup                            = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_NEWWORDSPERGROUP']
        g_strLblOldWordsPerGroup                            = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_OLDWORDSPERGROUP']
        g_strLblReviewPolicy                                = self.m_dictLanguages[g_strDefaultLang]['STR_LBL_REVIEWPOLICY']
        g_strCboReviewPolicy_item_LearnFirst                = self.m_dictLanguages[g_strDefaultLang]['STR_CBO_REVIEWPOLICY_ITEM_LEARNFIRST']
        g_strCboReviewPolicy_item_Random                    = self.m_dictLanguages[g_strDefaultLang]['STR_CBO_REVIEWPOLICY_ITEM_RANDOM']
        g_strCboReviewPolicy_item_ReviewFirst               = self.m_dictLanguages[g_strDefaultLang]['STR_CBO_REVIEWPOLICY_ITEM_REVIEWFIRST']
        
    
    def FreshCards(self):
        while len(self.boxCards_boxLibraryBody.children) != 0:
            self.boxCards_boxLibraryBody.remove(self.boxCards_boxLibraryBody.children[0])
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
            
            boxCardOperation = toga.Box(style = Pack(direction = COLUMN))
            btnEditCard = toga.Button("🔧", on_press = partial(self.cbEditExistedCard, eachCard))
            btnDeleteCard = toga.Button("❌", on_press = partial(self.cbDeleteCard, eachCard))
            boxCardOperation.add(
                btnEditCard,
                btnDeleteCard
            )
            
            boxCard.add(
                mtxtCardInfo,
                boxCardOperation
            )
            self.boxCards_boxLibraryBody.add(boxCard)
    
    def FreshWordLists(self):
        while len(self.boxWordLists_boxLibraryBody.children) != 0:
            self.boxWordLists_boxLibraryBody.remove(self.boxWordLists_boxLibraryBody.children[0])
        lsAllCardTypes = [""]
        for eachCard in self.m_dictTmpCards:
            lsAllCardTypes.append(eachCard)
        for eachWordList in self.m_dictTmpWordLists:
            boxWordList = toga.Box(style = Pack(direction = COLUMN))

            boxWordListInfo = toga.Box(style = Pack(direction = ROW))
            txtWordListInfo = toga.TextInput(style = Pack(flex = 1), readonly = True)

            strWordListInfo = eachWordList + f" {g_strFilePath}:" + self.m_dictTmpWordLists[eachWordList]['FilePath']
            txtWordListInfo.value = strWordListInfo
            btnEditWordList = toga.Button("🔧", on_press = partial(self.cbEditExistedWordList, eachWordList))
            btnDeleteWordList = toga.Button("❌", on_press = self.cbDeleteWordList, id = eachWordList)
            boxWordListInfo.add(
                txtWordListInfo,
                btnEditWordList,
                btnDeleteWordList
            )

            boxWordListConfig = toga.Box(style = Pack(direction = COLUMN))

            boxWordListCards = toga.Box(style = Pack(direction = ROW, alignment = CENTER))

            lblCard1 = toga.Label(g_strCard1)
            cboCard1 = toga.Selection(style = Pack(flex = 1), items = lsAllCardTypes)
            cboCard1.value = self.m_dictTmpWordLists[eachWordList]["CardType1"]
            cboCard1.on_select = partial(self.cbChangeWordListCardType, eachWordList, 1)

            lblCard2 = toga.Label(g_strCard2)
            cboCard2 = toga.Selection(style = Pack(flex = 1), items = lsAllCardTypes)
            cboCard2.value = self.m_dictTmpWordLists[eachWordList]["CardType2"]
            cboCard2.on_select = partial(self.cbChangeWordListCardType, eachWordList, 2)

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
            self.boxWordLists_boxLibraryBody.add(boxWordList)


    def BackToLibraryAndFresh(self):
        while len(self.boxBody.children) != 0:
            self.boxBody.remove(self.boxBody.children[0])
        self.btnIndex_boxNavigateBar.style.update(font_weight = NORMAL, background_color = g_clrBg, color = 'black')
        self.btnSettings_boxNavigateBar.style.update(font_weight = NORMAL, background_color = g_clrBg, color = 'black')
        self.btnLibrary_boxNavigateBar.style.update(font_weight = BOLD, background_color = g_clrPressedNavigationBtn, color = '#1651AA')
        self.boxBody.add(
            self.boxLibraryBody
        )
        self.FreshCards()
        self.FreshWordLists()
    
    def ChangeWordLearningBtns(self, iStatus):
        if iStatus == 0: # Clear all buttons
            while len(self.boxIndexBody_Row3Blank.children) != 0:
                self.boxIndexBody_Row3Blank.remove(self.boxIndexBody_Row3Blank.children[0])
            while len(self.boxIndexBody_Row5.children) != 0:
                self.boxIndexBody_Row5.remove(self.boxIndexBody_Row5.children[0])
        elif iStatus == 1: # Show the front
            self.boxIndexBody_Row5.add(
                self.btnShowBack
            )
        elif iStatus == 2: # Show the back
            self.boxIndexBody_Row3Blank.add(
                self.btnDelete
            )
            self.boxIndexBody_Row5.add(
                self.btnStrange,
                self.btnVague,
                self.btnFamiliar
            )

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

    def SaveWordListsAndCardsToFiles(self):
        global g_dictWordLists, g_dictCards
        jsonWordLists = json.dumps(g_dictWordLists)
        jsonCards = json.dumps(g_dictCards)
        with open(g_strWordlistsPath, 'w') as f:
            f.write(jsonWordLists)
        with open(g_strCardsPath, 'w') as f:
            f.write(jsonCards)
        print("-------------------------------------",g_strWordlistsPath)
        print(jsonWordLists)
        print("--------------------------------------",g_strCardsPath)
        print(jsonCards)
        with open(g_strWordlistsPath, 'r') as f:
            print(jsonWordLists)
        with open(g_strCardsPath, 'r') as f:
            print(jsonCards)

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

def main():
    return wordgets()
