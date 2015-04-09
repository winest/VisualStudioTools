function TraverseAndRelocateVcProjOutputIntermediate( aVcProjFolderPath , aOutputDir , aIntermediateDir , aLogFolder )
{
    var folder = WshFileSystem.GetFolder( aVcProjFolderPath );
    var enumFolder = new Enumerator( folder.SubFolders );
    for ( ; ! enumFolder.atEnd() ; enumFolder.moveNext() )
    {
        //Avoid the Backup folder
        if ( null == enumFolder.item().Name.match(/^Backup[0-9]*$/) )
        {
            TraverseAndRelocateVcProjOutputIntermediate( enumFolder.item().Path , aOutputDir , aIntermediateDir , aLogFolder );
        }
    }

    CWUtils.DbgMsg( "VERB" , "RelocateVcProjOutputIntermediate" , "Checking vcxproj files in \"" + aVcProjFolderPath + "\"" , aLogFolder );
    var enumFile = new Enumerator( folder.Files );
    for ( ; ! enumFile.atEnd() ; enumFile.moveNext() )
    {
        if ( enumFile.item().Name.match(/.+\.vcxproj/) )
        {
            CWUtils.DbgMsg( "VERB" , "RelocateVcProjOutputIntermediate" , "Updating " + enumFile.item().Name , aLogFolder );
            RelocateVcProjOutputIntermediate( enumFile.item().Path , aOutputDir , aIntermediateDir , aLogFolder );
        }
    }
}

function RelocateVcProjOutputIntermediate( aProjPath , aOutputDir , aIntermediateDir , aLogFolder )
{    
    CWUtils.DbgMsg( "VERB" , "RelocateVcProjOutputIntermediate" , "Enter. aProjPath=\"" + aProjPath + "\"" , aLogFolder );

    //Load file
    var fileProj = new CWUtils.CAdoTextFile();
    if ( false == fileProj.Open( aProjPath , "UTF-8" ) )
    {
        CWUtils.DbgMsg( "ERRO" , "RelocateVcProjOutputIntermediate" , "fileProj.Open() failed. aProjPath=" + aProjPath , aLogFolder );
        return false;
    }
    var strContent = fileProj.ReadAll();
    fileProj.Close();

    //Parse file into XML
    var xml = new CWUtils.CXml();
    if ( false == xml.LoadFromString(strContent) || 0 != xml.objDom.parseError.errorCode )
    {
        CWUtils.DbgMsg( "ERRO" , "RelocateVcProjOutputIntermediate" , "XML failed to parse with error " + 
                xml.objDom.parseError.errorCode + "(" + xml.objDom.parseError.reason + ")" , aLogFolder );
        return false;
    }
    else
    {
        //If the XML has xmlns, you must assign a prefix to access the namespace with XML 1.0
        var strSetNamespace = "xmlns:mynamespace=\"" + xml.GetNamespacesByName("Project")[0] + "\"";
        xml.document.setProperty( "SelectionNamespaces" , strSetNamespace );
    }
    
    //Update the variable in the OutDir tag
    if ( "undefined" !== typeof(aOutputDir) )
    {
        var aryResults = xml.GetElementsByTagName( "OutDir" );    
        for ( var i = 0 ; i < aryResults.length ; i++ )
        {
            CWUtils.DbgMsg( "INFO" , "RelocateVcProjOutputIntermediate" , aProjPath + " updating OutDir. " + aryResults[i].text + " => " + aOutputDir , aLogFolder );
            aryResults[i].text = aOutputDir;
        }
    }
        
    //Update the variable in the IntDir tag
    if ( "undefined" !== typeof(aIntermediateDir) )
    {
        var aryResults = xml.GetElementsByTagName( "IntDir" );    
        for ( var i = 0 ; i < aryResults.length ; i++ )
        {
            CWUtils.DbgMsg( "INFO" , "RelocateVcProjOutputIntermediate" , aProjPath + " updating IntDir. " + aryResults[i].text + " => " + aIntermediateDir , aLogFolder );
            aryResults[i].text = aIntermediateDir;
        }
    }

    //Overwrite the current project file
    xml.SaveToFile( aProjPath );

    CWUtils.DbgMsg( "VERB" , "RelocateVcProjOutputIntermediate" , "Leave" , aLogFolder );
    return true;
}
