option explicit



sub main()

    CreateProcess_

end sub

private sub CreateProcess_() ' {

   dim secAttrPrc as SECURITY_ATTRIBUTES ': secAttrPrc.nLength = len(secAttrPrc)
   dim secAttrThr as SECURITY_ATTRIBUTES ': secAttrThr.nLength = len(secAttrThr)

   dim startInfo  as STARTUPINFO ' : startInfo.cb = len(startInfo)

   dim procInfo   as PROCESS_INFORMATION

'       dwCreationFlags        :=   DEBUG_ONLY_THIS_PROCESS or CREATE_NEW_CONSOLE             , _
'       lpProcessAttributes    :=   secAttrPrc                                                , _
'       lpThreadAttributes     :=   secAttrThr                                                , _

   if CreateProcess (                                                                           _
        lpApplicationName      :=   vbNullString                                                  , _
        lpCommandLine          :=  "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE"   , _
        lpProcessAttributes    :=   nothing                                                  , _
        lpThreadAttributes     :=   nothing                                                  , _
        bInheritHandles        :=   false                                                     , _
        dwCreationFlags        :=   0                                                         , _
        lpEnvironment          :=   vbNullString                                              , _
        lpCurrentDirectory     :=   vbNullString                                              , _
        lpStartupInfo          :=   startInfo                                                 , _
        lpProcessInformation   :=   procInfo )  then

'      MsgBox "process created"

    else
       MsgBox GetLastErrorText
    end if

end sub ' }
