Function Get-NotPermissions ($path,$group,$error_rights,$error_long) {
#--- очищаем лог ошибок
    $error.Clear |out-null
#--- рекурсивно перебираем все папки
    foreach ($item in Get-ChildItem -LiteralPath $path -Recurse -Force -ErrorAction  SilentlyContinue -ErrorVariable error_mass | Where-Object {$_.PSIsContainer}) { 
#--- создаем массив с разрешениями
        $groups = ($item.PSPath |get-acl).Access | select -expandproperty IdentityReference
#--- проверяем есть ли в массиве интересующая нас группа
        if ($groups -notcontains $group) { '"'+$item.fullname+'"'}
    }
    
#--- проверяем существует ли файл и если чего удаляем его
    if (Test-Path $error_rights) {Remove-Item -Path $error_rights -Force -Confirm}
    if (Test-Path $error_long) {Remove-Item -Path $error_long -Force -Confirm}
    
#--- обрабатываем полученные ошибки
    foreach ($item in $error_mass ) {
#--- нету доступа
        if ($item.CategoryInfo.Reason -eq 'UnauthorizedAccessException') {
            '"' + $item.TargetObject + '"'  |Out-File -Encoding 'Unicode' -FilePath $error_rights -Append
        }
#--- длинный путь, больше 256 символов
        elseif ($item.CategoryInfo.Reason -eq 'PathTooLongException') {
            '"' + $item.TargetObject + '"'  |Out-File -Encoding 'Unicode' -FilePath $error_long -Append
        }
    }
}