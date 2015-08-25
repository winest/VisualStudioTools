@set @_PackJsInBatByWinest=0 /*
@ECHO OFF
CD /D "%~dp0"
CSCRIPT "%~0" //D //Nologo //E:JScript %1 %2 %3 %4 %5 %6 %7 %8 %9
IF %ERRORLEVEL% LSS 0 ( ECHO Failed. Error code is %ERRORLEVEL% )
PAUSE
EXIT /B
*/

var WshFileSystem = new ActiveXObject( "Scripting.FileSystemObject" );
var WshShell = WScript.CreateObject( "WScript.Shell" );
function LoadJs( aJsPath )
{
    var file = WshFileSystem.OpenTextFile( aJsPath , 1 );
    var strContent = file.ReadAll();
    file.Close();
    return strContent;
}
eval( LoadJs( "..\\..\\..\\_Include\\CWUtils\\JScript\\Windows\\CWFile.js" ) );
eval( LoadJs( "..\\..\\..\\_Include\\CWUtils\\JScript\\Windows\\CWStd.js" ) );
eval( LoadJs( "..\\..\\..\\_Include\\CWUtils\\JScript\\Windows\\CWShell.js" ) );
eval( LoadJs( "..\\..\\..\\_Include\\CWUtils\\JScript\\Windows\\CWXmlHttp.js" ) );
eval( LoadJs( "UpgradeVsSln.js" ) );
eval( LoadJs( "RelocateVcProjOutputIntermediate.js" ) );
eval( LoadJs( "RelocateCppProjInclude.js" ) );
eval( LoadJs( "RelocateCSharpProjInclude.js" ) );
eval( LoadJs( "BuildVsSln.js" ) );
eval( LoadJs( "RemoveVcGarbage.js" ) );


var strLogFolder = WshShell.CurrentDirectory + "\\Log";
CleanLogFolder();

for ( ;; )
{
    WScript.Echo( "\n\n\n========== Visual Studio Tools by winest ==========\n" );
    WScript.Echo( "What would you like to do?\n" +
                  "1. Upgrade all *.sln to the latest Visual Studio version\n" +
                  "2. Relocate the output and intermediate path in all *.vcxproj\n" +
                  "3. Relocate the relative path in all *.vcxproj\n" +
                  "4. Relocate the relative path in all *.csproj\n" +
                  "5. Build all *.sln\n" +
                  "6. Remove garbages generated by Visual Studio when compiling a user mode program\n" +
                  "7. Remove garbages generated by Visual Studio when compiling a kernel mode program\n" +
                  "8. Leave" );
    var strChoice = WScript.StdIn.ReadLine();
    switch ( strChoice )
    {
        case "1" :
        {
            var strVsPath = SelectVisualStudioPath();
            var strFolder = CWUtils.SelectFolder( "Please select the folder where *.sln exists. All *.sln files will be upgraded" );
            if ( true == CWUtils.SelectYesNo( "Upgrade *.sln files under \"" + strFolder + "\"? (y/n)" ) )
            {
                TraverseAndUpgradeVsSln( strVsPath , strFolder , strLogFolder );
            }
            break;
        }
        case "2" :
        {
            var strOutput = "$(SolutionDir)\\Output\\$(Configuration)\\$(Platform)\\";
            var strIntermediate = "$(SolutionDir)\\Intermediate\\$(Configuration)\\$(Platform)\\$(ProjectName)\\";
            var strFolder = CWUtils.SelectFolder( "Please select the folder where *.vcxproj exists. All *.vcxporj files will be updated" );
            if ( true == CWUtils.SelectYesNo( "Update *.vcxproj files under \"" + strFolder + "\"? (y/n)" ) )
            {
                TraverseAndRelocateVcProjOutputIntermediate( strFolder , strOutput , strIntermediate , strLogFolder );
            }
            break;
        }
        case "3" :
        {
            var aryFolderMap = [ new stFileMap( /.*(CWUtils|CWUtils\\Windows)$/ , "F:\\Code\\_Include\\CWUtils\\C++\\Windows" ,
                                                ["MyBuffer.cpp","MyBuffer.h","MyCmdArgsParser.cpp","MyCmdArgsParser.h","MyCrypto.cpp","MyCrypto.h","MyDllInjectClient.cpp","MyDllInjectClient.h","MyDllInjectCommonDef.h","MyDllInjectMgr.cpp","MyDllInjectMgr.h","MyDllInjectServer.cpp","MyDllInjectServer.h","MyEventMgr.cpp","MyEventMgr.h","MyFile.cpp","MyFile.h","MyFileMonitor.cpp","MyFileMonitor.h","MyGeneralUtils.cpp","MyGeneralUtils.h","MyHash.cpp","MyHash.h","MyHttpParser.cpp","MyHttpParser.h","MyIni.cpp","MyIni.h","MyLock.cpp","MyLock.h","MyMatrix.h","MyNetwork.cpp","MyNetwork.h","MyPlatform.cpp","MyPlatform.h","MyProcess.cpp","MyProcess.h","MyQueue.h","MyRegistry.cpp","MyRegistry.h","MySearch.cpp","MySearch.h","MyService.cpp","MyService.h","MyString.cpp","MyString.h","MyThreadPool.cpp","MyThreadPool.h","MyTime.cpp","MyTime.h","MyTree.h","MyVm.cpp","MyVm.h","MyVmWare.cpp","MyVmWare.h","MyVmWareBackdoor.h","MyVolume.cpp","MyVolume.h","MyWmiEventMonitor.cpp","MyWmiEventMonitor.h"] ,
                                                ["CWBuffer.cpp","CWBuffer.h","CWCmdArgsParser.cpp","CWCmdArgsParser.h","CWCrypto.cpp","CWCrypto.h","CWDllInjectClient.cpp","CWDllInjectClient.h","CWDllInjectCommonDef.h","CWDllInjectMgr.cpp","CWDllInjectMgr.h","CWDllInjectServer.cpp","CWDllInjectServer.h","CWEventMgr.cpp","CWEventMgr.h","CWFile.cpp","CWFile.h","CWFileMonitor.cpp","CWFileMonitor.h","CWGeneralUtils.cpp","CWGeneralUtils.h","CWHash.cpp","CWHash.h","CWHttpParser.cpp","CWHttpParser.h","CWIni.cpp","CWIni.h","CWLock.cpp","CWLock.h","CWMatrix.h","CWNetwork.cpp","CWNetwork.h","CWPlatform.cpp","CWPlatform.h","CWProcess.cpp","CWProcess.h","CWQueue.h","CWRegistry.cpp","CWRegistry.h","CWSearch.cpp","CWSearch.h","CWService.cpp","CWService.h","CWString.cpp","CWString.h","CWThreadPool.cpp","CWThreadPool.h","CWTime.cpp","CWTime.h","CWTree.h","CWVm.cpp","CWVm.h","CWVmWare.cpp","CWVmWare.h","CWVmWareBackdoor.h","CWVolume.cpp","CWVolume.h","CWWmiEventMonitor.cpp","CWWmiEventMonitor.h"]
                                              ) ,
                                 new stFileMap( /.*(Ui|CWUi\\Windows)$/ , "F:\\Code\\_Include\\CWUi\\C++\\Windows" , 
                                                ["MyButton.cpp","MyButton.h","MyCheckBox.cpp","MyCheckBox.h","MyComboBox.cpp","MyComboBox.h","MyCommonDlg.cpp","MyCommonDlg.h","MyControl.cpp","MyControl.h","MyDialog.cpp","MyDialog.h","MyEdit.cpp","MyEdit.h","MyLabel.cpp","MyLabel.h","MyListView.cpp","MyListView.h","MyRadioBtn.cpp","MyRadioBtn.h","MyToolbar.cpp","MyToolbar.h"] , 
                                                ["CWButton.cpp","CWButton.h","CWCheckBox.cpp","CWCheckBox.h","CWComboBox.cpp","CWComboBox.h","CWCommonDlg.cpp","CWCommonDlg.h","CWControl.cpp","CWControl.h","CWDialog.cpp","CWDialog.h","CWEdit.cpp","CWEdit.h","CWLabel.cpp","CWLabel.h","CWListView.cpp","CWListView.h","CWRadioBtn.cpp","CWRadioBtn.h","CWToolbar.cpp","CWToolbar.h"]
                                              )
                               ];
            var strFolder = CWUtils.SelectFolder( "Please select the folder where *.vcxproj exists. All *.vcxproj files will be updated" );
            if ( true == CWUtils.SelectYesNo( "Update *.vcxproj files under \"" + strFolder + "\"? (y/n)" ) )
            {
                TraverseAndRelocateCppProjInclude( strFolder , aryFolderMap , strLogFolder );
            }
            break;
        }
        case "4" :
        {
            var aryFolderMap = [ new stFileMap( /.*(CWUtils|..\\..\\..)$/ , "F:\\Code\\_Include\\CWUtils\\C#" ,
                                                [] ,
                                                []
                                              )
                               ];
            var strFolder = CWUtils.SelectFolder( "Please select the folder where *.csproj exists. All *.csproj files will be updated" );
            if ( true == CWUtils.SelectYesNo( "Update *.csproj files under \"" + strFolder + "\"? (y/n)" ) )
            {
                TraverseAndRelocateCSharpProjInclude( strFolder , aryFolderMap , strLogFolder );
            }
            break;
        }
        case "5" :
        {
            var strVsPath = SelectVisualStudioPath();
            var reBuildCfgs = SelectBuildCfgs();
            var strFolder = CWUtils.SelectFolder( "Please select the folder where *.sln exists. All *.sln files will be built" );
            if ( true == CWUtils.SelectYesNo( "Build *.sln files under \"" + strFolder + "\"? (y/n)" ) )
            {
                TraverseAndBuildVsSln( strVsPath , strFolder , reBuildCfgs , strLogFolder );
            }
            break;
        }
        case "6" :
        {
            var strFolder = CWUtils.SelectFolder( "Please select the folder where you want to clean" );
            if ( true == CWUtils.SelectYesNo( "Remove all user mode program garbages under \"" + strFolder + "\"? (y/n)" ) )
            {
                var reBlackFolders = /^(ipch|_UpgradeReport_Files[0-9]*|_tmh|Intermediate|Backup[0-9]*)$/;
                var reWhiteFiles = /^(TestWpp\.pdb)$/
                var reBlackFiles = /^(UpgradeLog[0-9]*\.htm|UpgradeLog[0-9]*\.XML|BuildLog[0-9]*\.htm)$/;

                //File extensions should use lower case
                var reBlackExts = /(\.obj|\.pdb|\.ilk|\.exp|\.map|\.ncb|\.sdf|\.pch|\.ipch|\.idb|\.dep|\.log|\.tlog|\.sbr|\.bsc|\.tmh|\.user|\.res|\.opt|\.plg|\.lastbuildstate|\.embed\.manifest|\.intermediate\.manifest)$/;
                RemoveVcGarbage( strFolder , reBlackFolders , reWhiteFiles , reBlackFiles , reBlackExts , strLogFolder );
            }
            break;
        }
        case "7" :
        {
            var strFolder = CWUtils.SelectFolder( "Please select the folder where you want to clean" );
            if ( true == CWUtils.SelectYesNo( "Remove all kernel mode program garbages under \"" + strFolder + "\"? (y/n)" ) )
            {
                var reBlackFolders = /^(ipch|_UpgradeReport_Files[0-9]*|_tmh|Intermediate|obj(chk|fre)_.+(x86|amd64))$/;
                var reWhiteFiles = /^(TestWpp\.pdb)$/
                var reBlackFiles = /^(UpgradeLog[0-9]*\.htm|UpgradeLog[0-9]*\.XML|BuildLog[0-9]*\.htm)$/;

                //File extensions should use lower case
                var reBlackExts = /(\.obj|\.ilk|\.exp|\.map|\.ncb|\.sdf|\.pch|\.ipch|\.idb|\.dep|\.wrn|\.err|\.log|\.tlog|\.sbr|\.bsc|\.tmh|\.user|\.res|\.opt|\.plg|\.lastbuildstate|\.embed\.manifest|\.intermediate\.manifest)$/;
                RemoveVcGarbage( strFolder , reBlackFolders , reWhiteFiles , reBlackFiles , reBlackExts , strLogFolder );
            }
            break;
        }
        case "8" :
        {
            if ( true == CWUtils.SelectYesNo( "Are you going to leave? (y/n)" ) )
            {
                WScript.Echo( "Successfully End" );
                WScript.Quit( 0 );
            }
            break;
        }
        default :
        {
            WScript.Echo( "Unknown choice: " + strChoice );
            break;
        }
    }
}


function CleanLogFolder()
{
    if ( WshFileSystem.FolderExists(strLogFolder) )
    {
        var folder = WshFileSystem.GetFolder( strLogFolder );

        var enumFolder = new Enumerator( folder.SubFolders );
        for ( ; ! enumFolder.atEnd() ; enumFolder.moveNext() )
        {
            WshFileSystem.DeleteFolder( enumFolder.item().Path , true );
        }
        var enumFile = new Enumerator( folder.Files );
        for ( ; ! enumFile.atEnd() ; enumFile.moveNext() )
        {
            WshFileSystem.DeleteFile( enumFile.item().Path , true );
        }
    }
}


function SelectVisualStudioPath()
{
    var strVsPath = CWUtils.ReadReg( "HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\Applications\\devenv.exe\\shell\\open\\command\\" );    
    var rePtn = /"(.+devenv\.exe)"*/;
    var aryInfo = rePtn.exec( strVsPath );
    if ( null != aryInfo )
    {
        strVsPath = aryInfo[1];
    }
    for ( ;; )
    {
        if ( null == strVsPath || false == WshFileSystem.FileExists(strVsPath) )
        {
            WScript.Echo( "Please enter the path of devenv.exe" );
            strVsPath = WScript.StdIn.ReadLine();
        }
        else
        {
            if ( true == CWUtils.SelectYesNo( "Is \"" + strVsPath + "\" ok for you? (y/n)" ) )
            {
                WScript.Echo( "Using \"" + strVsPath + "\"" );
                break;
            }
            else
            {
                strVsPath = null;
            }
        }
    }
    return strVsPath;
}

function SelectBuildCfgs()
{
    var aryChoices = [ "Debug\\|Win32" , "Debug\\|x64" , "Release\\|Win32" , "Release\\|x64" , "Debug\\|Any CPU" , "Release\\|Any CPU" ];
    var strBuildCfgs;
    for ( ;; )
    {
        WScript.Echo( "Please enter the config you want to build:" );
        for ( var i = 0 ; i < aryChoices.length ; i++ )
        {
            WScript.Echo( (i + 1) + ". " + aryChoices[i] );
        }
        WScript.Echo( "You can choose multiple options by entering 1+2+3+4, for example" );
        strBuildCfgs = WScript.StdIn.ReadLine();
        var aryBuildCfgs = strBuildCfgs.split( "+" );

        strBuildCfgs = "";
        for ( var i = 0 ; i < aryBuildCfgs.length ; i++ )
        {
            var nChoice = parseInt( aryBuildCfgs[i] ) - 1;
            if ( -1 < nChoice && nChoice < aryChoices.length )
            {
                if ( 0 < strBuildCfgs.length )
                {
                    strBuildCfgs += "|";
                }
                strBuildCfgs += aryChoices[nChoice];
            }
        }
        if ( 0 < strBuildCfgs.length &&
             true == CWUtils.SelectYesNo( "Is \"" + strBuildCfgs + "\" ok for you? (y/n)" ) )
        {
            strBuildCfgs = "^(" + strBuildCfgs + ")$";
            WScript.Echo( "Using \"" + strBuildCfgs + "\"" );
            return new RegExp( strBuildCfgs );
        }
    }
}
