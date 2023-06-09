Attribute VB_Name = "Demo"
'===================================================================================================
' Automating Chromium-Based Browsers with Chrome Dev Protocol API and VBA
'---------------------------------------------------------------------------------------------------
' Author(s)   :
'       ChrisK23 (Code Project)
' Contributors:
'       Long Vh (long.hoang.vu@hsbc.com.sg)
' Last Update :
'       27/04/23 Long Vh: made many improvements with v2.5 to make methods even more intuitive.
'       07/06/22 Long Vh: corrected typos in comments + more examples
'       03/06/22 Long Vh: codes edited + notes added + added extensive comments for HSBC colleagues
' References  :
'       Microsoft Scripting Runtime
' Notes       :
'       The framework does not need a matching webdriver as this is not a webdriver-based API.
'       This module includes a few examples of automating browsers using CDP. For the
'       engine codes, refer to the class modules CDPBrowser, CDPCore, CDPElement, and CDPJConv
'       For original examples, refer to Chris' article on CodeProject:
'       https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA
'       For the latest update of the CDP Framework by Long Vh:
'       https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA
'===================================================================================================
 
 
Sub runEdge()
'------------------------------------------------------
' This is an example of how to use the browser classes
' This demo tries to access a webpage of a famous movie
' and retrieve its current view count.
'------------------------------------------------------
 
   'Start Browser
   'If no browser name is indicated, chrome is started by default.
   'Homepage has been disabled to speed up by default.
   'To skip cleaning active sessions, set cleanActive to False.
   'This will make browser starts faster but at the risk of pipe error if
   'there are other chrome instances already running.
   'If reAttach = False, .start will not automatically try to reattach
   'to previous instances open by CDP but will start a brand new instead.
    Dim objBrowser As New CDPBrowser
    Dim HTMLDoc As New HTMLDocument
    Dim CDPWebElement As New Collection
    
    Dim objTable As Object
    Dim objTable2 As Object
    
    'objBrowser.start "edge", cleanActive:=True, reAttach:=False, addArgs:="--new-window"
    objBrowser.start "edge", cleanActive:=True, reAttach:=True
 
   'Navigate and wait
   'If till argument is omitted, will by default wait until ReadyState = complete
    ' https://www.isograd-testingservices.com/EN/testdetails?sbj_fam_id=15
    'objBrowser.navigate "https://livingwaters.com/movie/the-atheist-delusion/", isInteractive
 
    objBrowser.navigate "https://wsso-support.web.boeing.com:2016/redirect.html?URL=http://desktopportal.web.boeing.com/Inventory/Grading.aspx", isComplete
    objBrowser.wait
    
    objBrowser.sleep 5
    
    HTMLDoc.body.innerHTML = objBrowser.jsEval("document.body.innerHTML;") ' objIE.document.body.innerHTML
    
    Set CDPWebElement = objBrowser.getElementsByName("ScannedSerial")
    objBrowser.wait
    
    CDPWebElement(1).value = "SomeAssetTagHere"
'    Debug.Print CDPWebElement.innerText = "SomeTag"
'    With HTMLDoc
'        Set objTable = .getElementsByTagName("body")
'        Set objTable2 = .getElementByID("sbj_or_sbj_fam_id")
'        Debug.Print objTable2(1).innerText
'    End With
'    Set objDoc = objBrowser
    'objDoc.all.tags('select').Item(0)
'    Set objSelect = objBrowser.jsEval("document.all.tags('select').Item(0);")
'    Set objListItems = objBrowser.getElementByID("sbj_or_sbj_fam_id")
    
'    Debug.Print objListItems.innerHTML
    
    ' sbj_or_sbj_fam_id
   'Get view count
    'viewCount = objBrowser.jsEval("document.evaluate(""//h3[contains(., 'Total Views')]/*[1]"", document).iterateNext().innerText")
    'objBrowser.jsEval "alert(""This free movie has already reached " & viewCount & " views! Wow!"")"
    
 
End Sub
 
 
Sub runHidden()
'---------------------------------------------------------------------------------
' Demonstrate background running of an automated session.
' This demo will try to open Google in the background, then search for an article
' of CodeProject and retrieve its vote count. Once done, it will prompt a message
' to display the browser window.
' It is recommended to make Immediate Window visible so that you can see the
' activity that is running in the background.
' To confirm the result, you can perform the following steps:
'   1. Go to Google.com
'   2. Type "automate edge vba" and click Search
'   3. Click on the first result to reach the CodeProject's article
'   4. The vote count is seen there.
'---------------------------------------------------------------------------------
 
    Dim chrome As New CDPBrowser
 
   'Start and hide
    chrome.start
    chrome.hide
 
   'Perform automation in the background
    chrome.navigate "https://google.com", isInteractive
    chrome.getElementByQuery("[name='q']").value = "automate edge vba"
    chrome.getElementByQuery("[name='q']").submit
    
   'Click the target result link
    chrome.getElementByXPath("//h3[text()='Automate Chrome / Edge using VBA']").click
    
   'Get the vote count only once the target element appears on screen
   'The onExists method is needed as this element appears after ReadyState = "complete"
    voteCount = chrome.getElementByID("ctl00_RateArticle_VoteCountNoHist").onExist.innerHTML
    
   'Confirm result and display
    userChoice = MsgBox("Automation completed. Current vote counts: " & voteCount & ". Do you want to see the window?", vbYesNo)
    If userChoice = vbYes Then chrome.show Else chrome.quit
    
End Sub
 
 
Sub runTabsAsOne()
'--------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' Similar to the runInstances example but this is with multiple tabs in
' the same instance instead.
'--------------------------------------------------------------------------
 
    Dim chrome As New CDPBrowser
    chrome.start reAttach:=False
    chrome.show
    
   'Automate Tabs
    chrome.url = "google.com"   'or [chrome.navigate "google.com"]
    chrome.newTab "sg.yahoo.com"
    chrome.newTab "bing.com"
 
   'Resize to complete
    chrome.show xywh:="0 20 1000 700"
 
End Sub
 
 
Sub runTabsAsMany()
'-------------------------------------------------------------------------------
' Demonstrate the automation of multiple tabs in a single browser instance.
' This is like having 3 automation instances running together like runInstances.
' However, each tab will have to share the same start settings, unlike
' the case of runInstances where each instance can be setup with a different
' settings to each other.
'-------------------------------------------------------------------------------
 
    Dim chrome As New CDPBrowser
    chrome.start reAttach:=False
    chrome.show
 
   'Create and assign tabs
    Dim tab1 As New CDPBrowser                   'The keyword "New" is a must
    Dim tab2 As New CDPBrowser
    Dim tab3 As New CDPBrowser
    Set tab1 = chrome                            'The first tab is open by default after .start
    Set tab2 = chrome.newTab(newWindow:=True)    'newWindow: open tab as a new window instead of a tab
    Set tab3 = chrome.newTab(newWindow:=True)
 
   'Automate each tabs
    tab1.navigate "google.com"
    tab2.navigate "sg.yahoo.com"
    tab3.navigate "bing.com"
 
   'Resize to complete
    tab1.show xywh:="0 10 1000 700"
    tab2.show xywh:="0 45 1000 700"
    tab3.show xywh:="0 90 1000 700"
 
End Sub
 
 
Sub runNewTab()
'--------------------------------------------------------------------------
' This example demonstrates:
' 1. The use of advanced arguments feature added by Long Vh to
'    allow the choice of additional settings for the automation pipe. See
'    https://peter.sh/experiments/chromium-command-line-switches/
' 2. The xPath technique to directly modify the current HTML element
'    so that it will behave in a new way that it was not so before.
' 3. The technique employed to integrate the new tab open spontaneously
'    by interaction with the webpage (instead of using .newTab) into the
'    automation pipe for further processing on the new tab.
'--------------------------------------------------------------------------
 
   'Init browser with custom arguments
    Dim chrome As New CDPBrowser
    chrome.start addArgs:="--disable-popup-blocking"    'The disable-popup-blocking argument is needed to allow opening link in a new tab
    chrome.show asMaximized
    
   'Perform standard google search
    chrome.navigate "https://google.com"
    chrome.getElementByQuery("[name='q']").value = "newstarget.com"
    chrome.getElementByQuery("[name='q']").submit
 
   'Google search result returns links that open in the same tab window
   'For this demonstration, we need to make it open in a new tab window instead
    Dim targetElement As CDPElement
    Set targetElement = chrome.getElementByXPath(".//a[contains(@href, 'https://www.newstarget.com/')]")
    targetElement.setAttribute "target", "_blank"   'Modify the element attribute to open in a new tab instead
    targetElement.click                             'Click the link, a new tab will be spontaneously open
 
   'Use getTabNew to quickly refer to the next newly open tab
    Dim targetTab As New CDPBrowser
    Set targetTab = chrome.getTab
    targetTab.wait
 
   'Feed the top news title for today
    firstTitle = targetTab.getElementByQuery("div[class='Headline']").innerText
    targetTab.closeTab
    MsgBox "Top popular headline for the day is """ & firstTitle & """."
 
End Sub
 
 
Sub runIFrame()
'--------------------------------------------------------------------------
' This example demonstrates the CDP Framework v2.5 getIFrame technique for
' accessing iFrame element intuitively, an improvement over 1.0:
' 1. The use of App Mode via appUrl argument of the .start method.
' 2. The use of getIframe to easily access iFrame elements on the web page.
' 3. Working with a complex web design where nested iFrames are employed.
'--------------------------------------------------------------------------
    
    Dim demoUrl As String
    demoUrl = "https://www.w3schools.com/html/tryit.asp?filename=tryhtml_iframe_height_width"
    
    Dim chrome As New CDPBrowser
    chrome.start appUrl:=demoUrl, reAttach:=False
    
    Dim iFrame1 As CDPElement
    Dim iFrame2 As CDPElement
    Set iFrame1 = chrome.getElementByID("iframeResult").getIFrame
    Set iFrame2 = iFrame1.getElementByQuery("iframe[title='Iframe Example']").getIFrame
    
    txt = iFrame2.getElementByQuery("h1").innerText
    MsgBox "Retrieved text from the iFrame: """ & txt & """"
    
End Sub
 
 
Sub getSnapShot()
'--------------------------------------------------------------------------
' This example demonstrates the great enhancements of 2.5 over 1.0:
' 1. The use of external JS library via AddJsLib method.
' 2. The integration of external JS library into VBA project seamlessly.
' 3. The use of html2canvas in VBA, a very useful external library for
'    highly customized screenshot downloading.
'--------------------------------------------------------------------------
 
    Dim demoUrl As String
    demoUrl = "https://www.google.com/search?q=1sgd+to+vnd"
    
    Dim chrome As New CDPBrowser
    chrome.start                'not App Mode as sometimes Chrome App Mode does not allow file downloading
    chrome.navigate demoUrl
    
   'Snap a portion of the page based on the element indicator
   'If the second argument is omitted, snapPage will snap the entire page
    Dim targetArea As CDPElement
    Set targetArea = chrome.getElementByQuery("div[class='I4v0Kc wlkW8 PZPZlf']")
    chrome.snapPage "todaySGDvVND", targetArea
 
End Sub
 
 
Sub fillReactForm()
'-------------------------------------------------------------------------
' This example demonstrates the power of 2.6 on working natively
' with React form fields, which are notoriously complex to automate
' due to the fact that React form uses its own internal event handlings.
' The demo aims to:
' 1. Fill in the name field on the page.
' 2. Press submit.
' 3. If the field input is recognized by React, alert will tell its value.
'-------------------------------------------------------------------------
 
    Dim demoUrl As String
    demoUrl = "https://cdpn.io/gaearon/fullpage/VmmPgp?anon=true&editors=0010&view="
    
    Dim chrome As New CDPBrowser
    chrome.start
    chrome.navigate demoUrl
        
   'Get the target fields
    Dim ip As CDPElement
    Dim sb As CDPElement
    Set ip = chrome.getElementByID("result").getIFrame.getElementByQuery("input[type='text']")
    Set sb = chrome.getElementByID("result").getIFrame.getElementByQuery("input[type='submit']")
        
   'This traditional input method will fail as this is a React field
    chrome.jsEval ip.varName & ".value = 'TEST1'"
    chrome.jsEval ip.varName & ".dispatchEvent(new Event('input', { bubbles: true, simulated: true }))"
    sb.click 'you will not see "TEST1" in the alert result
 
   'This will succeed by using 2.6-enhanced .value property
    ip.value = "TEST2" '.value property is now overloaded with a smart React field detection & inputing
    sb.click
    
   'This will succeed as it mimicks sending raw keys but to a specific element
    ip.sendKeys "TEST3"
    sb.click
 
End Sub


Sub switchMain()
'---------------------------------------------------------------
' This example demonstrate the use of argument setMain to switch
' the main session tab to another tab so that future
' reattachment will hook this tab directly. This is useful if
' the main tab is supposed to be a tab open subsequently during
' the automation process by the target web link. The setMain
' method is preferrable to using "Set chrome = chrome.getTab..."
' because the latter method does not update the serial string
' for future reattachment.
'---------------------------------------------------------------
'    Dim objBrowser As New CDPBrowser
    Dim HTMLDoc As New HTMLDocument
    Dim CDPWebElement As New Collection
    
    Dim objTable As Object
    Dim objTable2 As Object


    Dim chrome As New CDPBrowser
    'chrome.start "edge", cleanActive:=True, reAttach:=True, addArgs:="--new-window"
    chrome.start "edge", cleanActive:=True, reAttach:=True
    chrome.newTab "google.com", setMain:=True   'the chrome object will now directly refer to the Google tab
    chrome.getTab("about:blank").closeTab       'prior 2.7, the next line will throw an error due to no main-switching mechanism
    chrome.printParams

    chrome.navigate "https://wsso-support.web.boeing.com:2016/redirect.html?URL=http://desktopportal.web.boeing.com/Inventory/Grading.aspx", isComplete
    chrome.wait
    
    chrome.sleep 5
    
    HTMLDoc.body.innerHTML = chrome.jsEval("document.body.innerHTML;") ' objIE.document.body.innerHTML
    
    Set CDPWebElement = chrome.getElementsByName("ScannedSerial")
    chrome.wait
    CDPWebElement(0).value = "NewAssetTagHere"
End Sub
