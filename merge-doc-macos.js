#!/usr/bin/env osascript -l JavaScript

//  merge-doc-macos.js
//
//  A port of TortoiseSVN/TortoiseGit merge-doc.js to to Open Scripting Architecture (OSA) on macOS.
//  For the source script, see https://github.com/TortoiseGit/TortoiseGit/blob/master/contrib/diff-scripts/merge-doc.js.
//
//  This file is distributed under the GNU General Public License.
//
//  Author: Zlatko Franjcic

// Microsoft Office versions for Microsoft Windows OS
const vOffice2000 = 9, vOffice2002 = 10, //vOffice2003 = 11,
    vOffice2007 = 12, vOffice2010 = 14, vOffice2013 = 15

// WdCompareTarget
const wdCompareTargetSelected = 'compare target selected'
//const wdCompareTargetCurrent = 'compare target current'
const wdCompareTargetNew = 'compare target new'
// WdViewType
const wdMasterView = 'master view'
const wdNormalView = 'normal view'
const wdOutlineView = 'outline view'
const wdReadingView = 'WordNote view' // 7

// WdSaveOptions
const wdDoNotSaveChanges = 'no'
//const wdPromptToSaveChanges = 'ask'
//const wdSaveChanges = 'yes'

// WdOpenFormat
const wdOpenFormatOpenFormatAuto = 'open format auto'

// MsoEncoding
const MsoEncodingUTF8 = 65001

// WdLineEndingType
const wdLineEndingTypeLineEndingCrLf = 'line ending cr lf'

function executeCompare(word, sBaseDoc, sOtherDoc, sOtherDocAuthor)
{
    var baseDoc, otherDoc, sTargetDoc
    
    var wdCompareTarget = (parseInt(word.version()) < vOffice2013 ? wdCompareTargetSelected : wdCompareTargetNew)
    
    // No 'activate' method -> comment code
    //baseDoc. activate //required otherwise it compares the wrong docs !!!
    // We cannot activate the document, so we open it, which should activate it
    baseDoc = word.open(null, {
        fileName: sBaseDoc,
        confirmConversions: true,
        readOnly: false,
        addToRecentFiles: false,
        repair: false,
        showingRepairs: false,
        passwordDocument: null,
        passwordTemplate: null,
        revert: false,
        writePassword: null,
        writePasswordTemplate: null,
        fileConverter: wdOpenFormatOpenFormatAuto
        })

    baseDoc.compare({
        path: sOtherDoc, 
        authorName: sOtherDocAuthor,
        target: wdCompareTargetNew,
        detectFormatChanges: true,
        ignoreAllComparisonWarnings: true,
        addToRecentFiles: false
    })

    if (parseInt(word.version()) < vOffice2013)
    {
        sTargetDoc = sOtherDoc;
    }
    else
    {
        // Due to sandboxing, we are going to save the new document to a temporary location that MS word has access to
        // This approach was inspired by: http://www.rondebruin.nl/mac/mac034.htm
        // Group container URL: https://developer.apple.com/library/content/documentation/Security/Conceptual/AppSandboxDesignGuide/AppSandboxInDepth/AppSandboxInDepth.html
        // The temporary file handling part was inspired by: http://nshipster.com/nstemporarydirectory/
        var groupContainerURL = $.NSFileManager.defaultManager.containerURLForSecurityApplicationGroupIdentifier('UBF8T346G9.Office')

        sTargetDoc = $.NSString.stringWithFormat('%@_%@', 
                    $.NSString.alloc.initWithUTF8String(sOtherDoc).lastPathComponent,
                    $.NSProcessInfo.processInfo.globallyUniqueString)
        sTargetDoc = groupContainerURL.URLByAppendingPathComponent(sTargetDoc).path.UTF8String
                
        word.activeDocument.saveAs({
            fileName: sTargetDoc,
            fileFormat: baseDoc.saveFormat,
            lockComments: false,
            password: null,
            addToRecentFiles: false,
            writePassword: null,
            readOnlyRecommended: false,
            embedTruetypeFonts: baseDoc.embedTrueTypeFonts,
            saveNativePictureFormat: false,
            saveFormsData: false,
            textEncoding: MsoEncodingUTF8,
            insertLineBreaks: false,
            allowSubstitutions: false,
            lineEndingType: wdLineEndingTypeLineEndingCrLf,
            htmlDisplayOnlyOutput: false,
            maintainCompatibility: true
        })
        
        // Close original document
        otherDoc = word.open(null, {
            fileName: sOtherDoc,
            confirmConversions: true,
            readOnly: false,
            addToRecentFiles: false,
            repair: false,
            showingRepairs: false,
            passwordDocument: null,
            passwordTemplate: null,
            revert: false,
            writePassword: null,
            writePasswordTemplate: null,
            fileConverter: wdOpenFormatOpenFormatAuto
            })

        otherDoc.close({saving: wdDoNotSaveChanges})
    }

    return sTargetDoc;
}

function executeMerge(word, baseDoc, sBaseDoc, sTheirDoc, sMyDoc)
{
    var theirDoc, myDocAfterCompare, sTheirDocAfterCompare, sMyDocAfterCompare
    
    theirDoc = baseDoc

    sTheirDocAfterCompare = executeCompare(word, sBaseDoc, sTheirDoc, 'theirs')

    sMyDocAfterCompare = executeCompare(word, sBaseDoc, sMyDoc, 'mine')
    
    //theirDoc.save({in:null, as: null})
    //myDoc.save({in:null, as: null})
    
    // No 'activate' method -> comment code
    //myDoc.activate //required? just in case
    // We cannot activate the document, so we open it, which should activate it
    myDocAfterCompare = word.open(null, {
        fileName: sMyDocAfterCompare,
        confirmConversions: true,
        readOnly: false,
        addToRecentFiles: false,
        repair: false,
        showingRepairs: false,
        passwordDocument: null,
        passwordTemplate: null,
        revert: false,
        writePassword: null,
        writePasswordTemplate: null,
        fileConverter: wdOpenFormatOpenFormatAuto
        })

    myDocAfterCompare.merge({fileName: sTheirDocAfterCompare})
    
    // Clean-up (this should work, even if the docs are still opened in Word)
    if (sTheirDoc != sTheirDocAfterCompare)
    {
        $.NSFileManager.defaultManager.removeItemAtPathError(sTheirDocAfterCompare, null)
    }

    if (sMyDoc != sMyDocAfterCompare)
    {
        $.NSFileManager.defaultManager.removeItemAtPathError(sMyDocAfterCompare, null)
    }
    
    // Built-in three-way merge does not work that nicely
    //myDoc.threeWayMerge({localDocument: myDoc, serverDocument: theirDoc, baseDocument: baseDoc, favorSource: false})
}

function run(argv)
{   
    ObjC.import('stdlib')
    ObjC.import('stdio')

    var word, sTheirDoc, sMyDoc, sBaseDoc, sMergedDoc, baseDoc

    argc = argv.length 
    if (argv.length < 4)
    {
        var scriptApp = Application.currentApplication()
        scriptApp.includeStandardAdditions = true
        const basename = $.NSString.alloc.initWithUTF8String(
                            scriptApp.pathTo(this))
                            .lastPathComponent.UTF8String
        $.printf('Usage: %s <absolute-path-to-merged.doc> <absolute-path-to-theirs.doc> <absolute-path-to-mine.doc> <absolute-path-to-base.doc>\n', basename)
        $.exit(1)
    }

    sMergedDoc = argv[0]
    sTheirDoc = argv[1]
    sMyDoc = argv[2]
    sBaseDoc = argv[3]
    
    if (!$.NSFileManager.defaultManager.fileExistsAtPath(sTheirDoc))
    {
        $.printf('File %s does not exist. Cannot compare the documents.\n', sTheirDoc)
        $.exit(1)
    }
    
    if (!$.NSFileManager.defaultManager.fileExistsAtPath(sMergedDoc))
    {
        $.printf('File %s does not exist. Cannot compare the documents.\n', sMergedDoc)
        $.exit(1)
    }

    try {
        word = Application('com.microsoft.Word')
    }
    catch(ex)
    {
        $.printf('You must have Microsoft Word installed to perform this operation.\n')
        $.exit(1)
    }
        
    // The 'visible' property does not exist in this interface
    //word.visible
    
    // Open the base document
    baseDoc = word.open(null, {
        fileName: sTheirDoc,
        confirmConversions: true,
        readOnly: false,
        addToRecentFiles: false,
        repair: false,
        showingRepairs: false,
        passwordDocument: null,
        passwordTemplate: null,
        revert: false,
        writePassword: null,
        writePasswordTemplate: null,
        fileConverter: wdOpenFormatOpenFormatAuto
        })

    try
    {
        // Merge into the "My" document
        if (parseInt(word.version()) < vOffice2000)
        {
            // Contrary to the original TortoiseSVN/Git script, we cannot use duck typing -> comment out this line,
            // as we only support the newer interface below
            //baseDoc.compare({path: sMergedDoc})
            $.printf('Warning: Office versions up to Office 2000 are not officially supported.\n');
            baseDoc.compare({
                path: sMergedDoc, 
                authorName: 'Comparison',
                target: wdCompareTargetNew,
                detectFormatChanges: true,
                ignoreAllComparisonWarnings: true,
                addToRecentFiles: false
            })
        }
        else if (parseInt(word.version()) < vOffice2007)
        {
            baseDoc.compare({
                path: sMergedDoc, 
                authorName: 'Comparison',
                target: wdCompareTargetNew,
                detectFormatChanges: true,
                ignoreAllComparisonWarnings: true,
                addToRecentFiles: false
            })
        }
        else if (parseInt(word.version()) < vOffice2010)
        {
            baseDoc.merge({fileName: sMergedDoc})
        }
        else
        {
             //2010 and later - handle slightly differently as the basic merge isn't that good
            //note this is designed specifically for svn 3 way merges, during the commit conflict resolution process
            executeMerge(word, baseDoc, sBaseDoc, sTheirDoc, sMyDoc)
        }
       
        // Show the merge result
        if (parseInt(word.version()) < vOffice2007)
        {
            word.activeDocument.windows[0].visible = true
        }

        // Close the first document
        if ((parseInt(word.version()) >= vOffice2002) && (parseInt(word.version()) < vOffice2010))
        {
            baseDoc.close({saving: wdDoNotSaveChanges})
        }
             
        // Show usage hint message
        try {
            var scriptApp = Application.currentApplication()
            scriptApp.includeStandardAdditions = true
        
            var sButtonOK = 'OK', sButtonCancel = 'Cancel'
            scriptApp.displayAlert('MacOS Word Merge', {
                message: 'You have to accept or reject the changes before\nsaving the document to prevent future problems.\n\nWould you like to see a help page on how to do this?',
                as: 'informational',
                buttons: [sButtonOK, sButtonCancel],
                defaultButton: sButtonOK,
                cancelButton: sButtonCancel
            })

            // If we get to here, then OK was clicked
            //var urlString = 'http://office.microsoft.com/en-us/assistance/HP030823691033.aspx' // URL found in original TSVN script 
            var urlString = 'https://support.office.com/en-us/article/Review-accept-reject-and-hide-tracked-changes-8af4088d-365f-4461-a75b-35c4fc7dbabd'
            scriptApp.openLocation(urlString)
        }
        catch(ex)
        {
            if (ex.errorNumber != -128)
            {
                $.printf('Failed to usage hint message dialog: %s\n', ex.message);
            }
        }
    }
    catch(ex)
    {
        $.printf('Error running merge (merged: %s, theirs: %s, mine: %s, base: %s): %s\n', 
            sMergedDoc,
            sTheirDoc, sMyDoc,
            sBaseDoc,
            ex.message)
        // Quit
        $.exit(1)
    }
    $.exit(0)
}