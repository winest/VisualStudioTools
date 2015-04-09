function TraverseAndBuildVsSln( aVsPath , aSlnFolderPath , aBuildCfgs , aLogFolder )
{
    if ( "undefined" === typeof(aBuildCfgs) || null == aBuildCfgs )
    {
        CWUtils.DbgMsg( "VERB" , "BuildVsSln" , "aBuildCfgs is empty" , aLogFolder );
        return false;
    }

    var folder = WshFileSystem.GetFolder( aSlnFolderPath );
    var enumFolder = new Enumerator( folder.SubFolders );
    for ( ; ! enumFolder.atEnd() ; enumFolder.moveNext() )
    {
        //Avoid build the *.sln files in the Backup folder
        if ( null == enumFolder.item().Name.match(/^Backup[0-9]*$/) )
        {
            TraverseAndBuildVsSln( aVsPath , enumFolder.item().Path , aBuildCfgs , aLogFolder );
        }
    }

    CWUtils.DbgMsg( "VERB" , "BuildVsSln" , "Checking sln files in \"" + aSlnFolderPath + "\"" , aLogFolder );
    var enumFile = new Enumerator( folder.Files );
    for ( ; ! enumFile.atEnd() ; enumFile.moveNext() )
    {
        if ( enumFile.item().Name.match(/.+\.sln/) )
        {
            CWUtils.DbgMsg( "VERB" , "BuildVsSln" , "Building " + enumFile.item().Name , aLogFolder );
            BuildVsSln( aVsPath , enumFile.item().Path , aBuildCfgs , aLogFolder );
        }
    }
}

function BuildVsSln( aVsPath , aSlnPath , aBuildCfgs , aLogFolder )
{
    if ( "undefined" === typeof(aBuildCfgs) || null == aBuildCfgs )
    {
        CWUtils.DbgMsg( "VERB" , "BuildVsSln" , "aBuildCfgs is empty" , aLogFolder );
        return false;
    }

    //Parse sln file to check what build types are supported
    var fileSln = new CWUtils.CAdoTextFile();
    if ( false == fileSln.Open( aSlnPath , "utf-8" ) )
    {
        CWUtils.DbgMsg( "ERRO" , "BuildVsSln" , "fileSln.Open() failed. aSlnPath=" + aSlnPath , aLogFolder );
        return false;
    }

    var bDetectSlnCfg = false;
    var arySlnCfgs = [];
    var reCfgs = /^\s*(.+)\s*=\s*(.+)$/
    while ( false == fileSln.AtEndOfStream() )
    {
        var strLine = fileSln.ReadLine();
        if ( -1 != strLine.indexOf("GlobalSection(SolutionConfigurationPlatforms)") )
        {
            bDetectSlnCfg = true;
            continue;
        }
        if ( -1 != strLine.indexOf("EndGlobalSection") )
        {
            bDetectSlnCfg = false;
            continue;
        }

        if ( bDetectSlnCfg )
        {
            var cfgs = reCfgs.exec( strLine );
            if ( 1 < cfgs.length )
            {
                arySlnCfgs.push( cfgs[1].replace(/^\s+/,"").replace(/\s+$/,"") );
            }
        }
    }

    //Compare the supported config with aBuildCfgs and build the solution if matched
    for ( var i = 0 ; i < arySlnCfgs.length ; i++ )
    {
        if ( arySlnCfgs[i].match(aBuildCfgs) )
        {
            var strCmdRebuild = "\"" + aVsPath + "\" \"" + aSlnPath + "\" /Rebuild \"" + arySlnCfgs[i] + "\"";
            var execObj = CWUtils.Exec( strCmdRebuild , true , null );
            if ( 0 != execObj.ExitCode )
            {
                CWUtils.DbgMsg( "ERRO" , "BuildVsSln" , "Failed with exit code " + uExitCode + " when doing " + strCmdRebuild , aLogFolder );
            }
            else
            {
                CWUtils.DbgMsg( "INFO" , "BuildVsSln" , "Succeed to do " + strCmdRebuild , aLogFolder );
            }
        }
    }
}
