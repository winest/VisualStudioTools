function RemoveVcGarbage( aFolderPath , aBlackFolders , aWhiteFiles , aBlackFiles , aBlackExts , aLogFolder )
{
    if ( "undefined" === typeof(aBlackFolders) )
    {
        aBlackFolders = null;
    }
    if ( "undefined" === typeof(aWhiteFiles) )
    {
        aWhiteFiles = null;
    }
    if ( "undefined" === typeof(aBlackFiles) )
    {
        aBlackFiles = null;
    }
    if ( "undefined" === typeof(aBlackExts) )
    {
        aBlackExts = null;
    }
    //CWUtils.DbgMsg( "INFO" , "RemoveVcGarbage" , "Enter. aFolderPath=" + aFolderPath + 
    //        "\taBlackFolders=" + aBlackFolders + "\taWhiteFiles=" + aWhiteFiles +
    //        "\taBlackFiles=" + aBlackFiles + "\taBlackExts=" + aBlackExts );
    var folder = WshFileSystem.GetFolder( aFolderPath );
    var enumSubFolder = new Enumerator( folder.SubFolders );    
    var uSubFolderCnt = 0;
    for ( ; ! enumSubFolder.atEnd() ; enumSubFolder.moveNext() )
    {
        uSubFolderCnt++;
        var strSubFolderName = enumSubFolder.item().Name;

        //Remove default intermediate folders if the current folder name equals its subfolder name
        if ( WshFileSystem.GetBaseName(aFolderPath) == strSubFolderName )
        {
            RemoveIntermediateFolders( aFolderPath , aLogFolder );
        }

        if ( strSubFolderName.match( aBlackFolders ) )
        {
            var strBlackFolder = aFolderPath + "\\" + strSubFolderName;

            CWUtils.DbgMsg( "INFO" , "RemoveVcGarbage" , "Deleting folder \"" + strBlackFolder + "\"" , aLogFolder );
            WshFileSystem.DeleteFolder( strBlackFolder , true );
            uSubFolderCnt--;
        }
        else
        {
            RemoveVcGarbage( aFolderPath + "\\" + strSubFolderName , aBlackFolders , aWhiteFiles , aBlackFiles , aBlackExts , aLogFolder );
        }
    }
    var enumFile = new Enumerator( folder.Files );
    var uFileCnt = 0;
    for ( ; ! enumFile.atEnd() ; enumFile.moveNext() )
    {
        uFileCnt++;
        var strFileName = enumFile.item().Name;
        if ( ( ! strFileName.match(aWhiteFiles) ) &&
             ( strFileName.match(aBlackFiles) || strFileName.toLowerCase().match(aBlackExts) ) )
        {
            CWUtils.DbgMsg( "INFO" , "RemoveVcGarbage" , "Deleting file \"" + aFolderPath + "\\" + strFileName + "\"" , aLogFolder );
            WshFileSystem.DeleteFile( aFolderPath + "\\" + strFileName , true );
            uFileCnt--;
        }
    }
    
    if ( 0 == uSubFolderCnt && 0 == uFileCnt )
    {
        CWUtils.DbgMsg( "INFO" , "RemoveVcGarbage" , "Deleting empty folder \"" + aFolderPath + "\"" , aLogFolder );
        WshFileSystem.DeleteFolder( aFolderPath , true );
    }
}

function RemoveIntermediateFolders( aFolderPath , aLogFolder )
{
    var strFolderName = WshFileSystem.GetBaseName( aFolderPath );
    var blackFolders = new Array();
    blackFolders.push( aFolderPath + "\\" + strFolderName + "\\Debug" );
    blackFolders.push( aFolderPath + "\\" + strFolderName + "\\Release" );
    blackFolders.push( aFolderPath + "\\" + strFolderName + "\\Win32" );
    blackFolders.push( aFolderPath + "\\" + strFolderName + "\\x64" );
    
    for ( var i = 0 ; i < blackFolders.length ; i++ )
    {
        if ( WshFileSystem.FolderExists( blackFolders[i] ) )
        {
            CWUtils.DbgMsg( "INFO" , "RemoveVcGarbage" , "Deleting intermediate folder \"" + blackFolders[i] + "\"" , aLogFolder );
            WshFileSystem.DeleteFolder( blackFolders[i] , true );
        }
    }
}
