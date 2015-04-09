function TraverseAndUpgradeVsSln( aVsPath , aSlnFolderPath , aLogFolder )
{
    var folder = WshFileSystem.GetFolder( aSlnFolderPath );
    var enumFolder = new Enumerator( folder.SubFolders );
    for ( ; ! enumFolder.atEnd() ; enumFolder.moveNext() )
    {
        //Avoid update the *.sln files in the Backup folder. It might cause infinite loop
        if ( null == enumFolder.item().Name.match(/^Backup[0-9]*$/) )
        {
            TraverseAndUpgradeVsSln( aVsPath , enumFolder.item().Path , aLogFolder );
        }
    }

    CWUtils.DbgMsg( "VERB" , "UpgradeVsSln" , "Checking sln files in \"" + aSlnFolderPath + "\"" , aLogFolder );
    var enumFile = new Enumerator( folder.Files );
    for ( ; ! enumFile.atEnd() ; enumFile.moveNext() )
    {
        if ( enumFile.item().Name.match(/.+\.sln/) )
        {
            CWUtils.DbgMsg( "VERB" , "UpgradeVsSln" , "Upgrading " + enumFile.item().Name , aLogFolder );
            UpgradeVsSln( aVsPath , enumFile.item().Path , aLogFolder );
        }
    }
}

function UpgradeVsSln( aVsPath , aSlnPath , aLogFolder )
{
    var strCmdUpgrade = "\"" + aVsPath + "\" \"" + aSlnPath + "\" /Upgrade";
    var execObj = CWUtils.Exec( strCmdUpgrade , true , null );
    if ( 0 != execObj.ExitCode )
    {
        CWUtils.DbgMsg( "ERRO" , "UpgradeVsSln" , "Failed with exit code " + execObj.ExitCode + " when doing " + strCmdUpgrade , aLogFolder );
    }
    else
    {
        CWUtils.DbgMsg( "INFO" , "UpgradeVsSln" , "Succeed to do " + strCmdUpgrade , aLogFolder );
    }
}
