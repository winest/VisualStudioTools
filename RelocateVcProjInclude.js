function stFileMap( aOldFolder , aNewFolder , aOldFiles , aNewFiles )
{
    this.reOldFolder = new RegExp( aOldFolder );
    this.strNewFolder = aNewFolder;

    //The length of aryOldFiles and aryNewFiles must be the same
    this.aryOldFiles = aOldFiles;
    this.aryNewFiles = aNewFiles;
}

function TraverseAndRelocateVcProjInclude( aVcProjFolderPath , aFileMapAry , aLogFolder )
{
    var folder = WshFileSystem.GetFolder( aVcProjFolderPath );
    var enumFolder = new Enumerator( folder.SubFolders );
    for ( ; ! enumFolder.atEnd() ; enumFolder.moveNext() )
    {
        //Avoid the Backup folder
        if ( null == enumFolder.item().Name.match(/^Backup[0-9]*$/) )
        {
            TraverseAndRelocateVcProjInclude( enumFolder.item().Path , aFileMapAry , aLogFolder );
        }
    }

    CWUtils.DbgMsg( "VERB" , "RelocateVcProjInclude" , "Checking vcxproj files in \"" + aVcProjFolderPath + "\"" , aLogFolder );
    var enumFile = new Enumerator( folder.Files );
    for ( ; ! enumFile.atEnd() ; enumFile.moveNext() )
    {
        if ( enumFile.item().Name.match(/.+\.vcxproj/) )
        {
            CWUtils.DbgMsg( "VERB" , "RelocateVcProjInclude" , "Updating " + enumFile.item().Name , aLogFolder );
            RelocateVcProjInclude( enumFile.item().Path , aFileMapAry , aLogFolder );
        }
    }
}

function RelocateVcProjInclude( aProjPath , aFileMapAry , aLogFolder )
{    
    CWUtils.DbgMsg( "VERB" , "RelocateVcProjInclude" , "Enter. aProjPath=\"" + aProjPath + "\"" , aLogFolder );

    //Load file
    var fileProj = new CWUtils.CAdoTextFile();
    if ( false == fileProj.Open( aProjPath , "UTF-8" ) )
    {
        CWUtils.DbgMsg( "ERRO" , "RelocateVcProjInclude" , "fileProj.Open() failed. aProjPath=" + aProjPath , aLogFolder );
        return false;
    }
    var strContent = fileProj.ReadAll();
    fileProj.Close();

    //Parse file into XML
    var xml = new CWUtils.CXml();
    if ( false == xml.LoadFromString(strContent) || 0 != xml.objDom.parseError.errorCode )
    {
        CWUtils.DbgMsg( "ERRO" , "RelocateVcProjInclude" , "XML failed to parse with error " + 
                xml.objDom.parseError.errorCode + "(" + xml.objDom.parseError.reason + ")" , aLogFolder );
        return false;
    }
    else
    {
        //If the XML has xmlns, you must assign a prefix to access the namespace with XML 1.0
        var strSetNamespace = "xmlns:mynamespace=\"" + xml.GetNamespacesByName("Project")[0] + "\"";
        xml.document.setProperty( "SelectionNamespaces" , strSetNamespace );
    }
    
    //Update the variable in the AdditionalIncludeDirectories tag
    var aryResults = xml.GetElementsByTagName( "AdditionalIncludeDirectories" );
    for ( var i = 0 ; i < aryResults.length ; i++ )
    {
        var strNewText = "";
        var aryAdditionalIncludeDirectories = aryResults[i].text.split( ";" );
        for ( var j = 0 ; j < aryAdditionalIncludeDirectories.length ; j++ )
        {
            if ( 0 < strNewText.length )
            {
                strNewText += ";"
            }
            
            var bMatched = false;
            for ( var k = 0 ; k < aFileMapAry.length ; k++ )
            {
                if ( "undefined" !== typeof(aFileMapAry[k].reOldFolder) && 
                     aryAdditionalIncludeDirectories[j].match(aFileMapAry[k].reOldFolder) )
                {
                    strNewText += CWUtils.ComputeRelativePath( aProjPath , aFileMapAry[k].strNewFolder );
                    bMatched = true;
                    break;
                }
            }
            if ( false == bMatched )
            {
                strNewText += aryAdditionalIncludeDirectories[j];
            }
        }

        CWUtils.DbgMsg( "INFO" , "RelocateVcProjInclude" , aProjPath + " updating AdditionalIncludeDirectories. " + aryResults[i].text + " => " + strNewText , aLogFolder );
        aryResults[i].text = strNewText;
    }
    
    //Update the variable in the ClCompile.Include and ClInclude.Include
    var aryInclude = xml.document.selectNodes("//mynamespace:ClCompile[@Include] | //mynamespace:ClInclude[@Include]");
    for ( var i = 0 ; i < aryInclude.length ; i++ )
    {
        var strNewPath;
        var bMatched = false;
        var bEverSearchInFolder = false;
        var strInclude = aryInclude[i].getAttribute( "Include" );
        var strFolder = WshFileSystem.GetParentFolderName( strInclude );
        var strFile = WshFileSystem.GetFileName( strInclude );

        for ( var j = 0 ; j < aFileMapAry.length ; j++ )
        {
            //Remap files if the folder name match reOldFolder and the same filename or strNewFiles is found under strNewFolder
            if ( 0 < aFileMapAry[j].strNewFolder.length &&
                 '\\' != aFileMapAry[j].strNewFolder.charAt(aFileMapAry[j].strNewFolder.length - 1))
            {
                aFileMapAry[j].strNewFolder += "\\";
            }

            if ( "undefined" !== aFileMapAry[j].reOldFolder &&
                 strFolder.match(aFileMapAry[j].reOldFolder) )
            {
                bEverSearchInFolder = true;

                if ( WshFileSystem.FileExists( aFileMapAry[j].strNewFolder + strFile ) )
                {
                    strNewPath = CWUtils.ComputeRelativePath( aProjPath , aFileMapAry[j].strNewFolder + strFile );
                    bMatched = true;
                }
                else
                {
                    for ( var k = 0 ; k < aFileMapAry[j].aryOldFiles.length ; k++ )
                    {
                        if ( "undefined" !== aFileMapAry[j].aryOldFiles[k] &&
                             "undefined" !== aFileMapAry[j].aryNewFiles[k] &&
                             aFileMapAry[j].aryOldFiles.length == aFileMapAry[j].aryNewFiles.length &&    
                             strFile.match(aFileMapAry[j].aryOldFiles[k]) &&
                             WshFileSystem.FileExists(aFileMapAry[j].strNewFolder + aFileMapAry[j].aryNewFiles[k]) )
                        {
                            strNewPath = CWUtils.ComputeRelativePath( aProjPath , aFileMapAry[j].strNewFolder + aFileMapAry[j].aryNewFiles[k] );
                            bMatched = true;
                            break;
                        }
                    }
                }

                if ( true == bMatched )
                {
                    break;
                }
            }
        }
        if ( false == bMatched )
        {
            if ( true == bEverSearchInFolder )
            {
                CWUtils.DbgMsg( "ERRO" , "RelocateVcProjInclude" , aProjPath + " updating Include. " + strFile + " is not found under any new folder" , aLogFolder );
            }
            strNewPath = strInclude;
        }

        CWUtils.DbgMsg( "INFO" , "RelocateVcProjInclude" , aProjPath + " updating Include. " + strInclude + " => " + strNewPath , aLogFolder );
        aryInclude[i].setAttribute( "Include" , strNewPath );
    }

    //Overwrite the current project file
    xml.SaveToFile( aProjPath );

    CWUtils.DbgMsg( "VERB" , "RelocateVcProjInclude" , "Leave" , aLogFolder );
    return true;
}
