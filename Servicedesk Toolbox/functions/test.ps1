$path = "\\LABDC01\FileShare"
    
try{
    $accessRights = (Get-ACL $path).access | select IdentityReference,FileSystemRights,AccessControlType,IsInherited,InheritanceFlags
}catch{
    Write-Host $error
    $error.Clear()
    return
}


foreach($access in $accessRights){
    Write-Host $access
}
