Dim lib : Set lib = New LibLoader
lib.path = "../../lib"
lib.Import "cactus.vbs"



Call Main

Sub Main()
    Dim obj_vc_project_doc : Set obj_vc_project_doc = New vc_project_doc
    obj_vc_project_doc.filename = WScript.Arguments(0)
    'obj_vc_project_doc.filename = "zlib.vcxproj"
    obj_vc_project_doc.parse()
    obj_vc_project_doc.export_vc2005_doc()
    Set obj_vc_project_doc = Nothing
End Sub
	

Class vc_project_doc
    Private filename_
    Private doc_
    Private root_      
   
    
    
    Private config_type     
   
   
    
    Private preprocessor_definitions_               ' 预定义        
    Private additional_include_directories_         ' include 目录
    Private resfiles, resfiles2
    Private hfiles
    Private cfiles

    Private Sub Class_Initialize()          
    End Sub  
  
    Private Sub Class_Terminate()      
        Set doc_ = Nothing
    End Sub  

    Public Property Let filename(value)   
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        filename_ = fso.GetAbsolutePathName(value)
        Set fso = Nothing          
        
        Set doc_ = CreateObject("MSXML2.DOMDocument")
        doc_.async = False
        doc_.load(filename_)
        If doc_.parseError.errorCode = 0 Then
            Set root_ = doc_.documentElement		
        End If
    End Property
    
    Public Property Get filename()
        filename = filename_
    End Property    

    Private Function get_project_guid()
        Dim node
        Set node = root_.selectSingleNode("//ProjectGuid")
        If Not (node is Nothing) Then
            get_project_guid = node.text
        End If 
    End Function

    Private Function get_keyword()
        Dim node
        Set node = root_.selectSingleNode("//Keyword")
        If Not (node is Nothing) Then
            get_keyword = node.text
        End If 
    End Function

    Private Function get_root_namespace()
        Dim node
        Set node = root_.selectSingleNode("//root_Namespace")
        If Not (node is Nothing) Then
            get_root_namespace = node.text
        End If 
    End Function

    Private Function get_xpath_by_type(xpath, ByVal type_mode)
        If type_mode = "Debug" Then
            get_xpath_by_type = Replace(xpath, "Release", "Debug")
        Else
            get_xpath_by_type = xpath
        End If        
    End Function

    Private Function get_config_type(type_mode)
        Dim node
        Set node = root_.selectSingleNode(get_xpath_by_type("//PropertyGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""]/ConfigurationType", type_mode))
        If Not (node is Nothing) Then   
            get_config_type = node.text
        End If
    End Function

    Private Function get_precompiled_header(type_mode)
        Dim node
        Set node = root_.selectSingleNode(get_xpath_by_type("//ItemDefinitionGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""]//ClCompile/PrecompiledHeader", type_mode))
        If (Not node is Nothing) Then
            If node.text = "NotUsing" Then
                get_precompiled_header = "0"
            End If
        Else
            get_precompiled_header = "0"
        End If
    End Function

    Private Function get_output_file(type_mode)
        If type_mode = "Debug" Then
            get_output_file = "../../product/win32d/$(TargetName)$(TargetExt)"
        Else
            get_output_file = "../../product/win32/$(TargetName)$(TargetExt)"
        End If
    End Function

    Private Function get_output_lib_dir(type_mode)
        If type_mode = "Debug" Then
            get_output_lib_dir = "../../lib/win32d"
        Else
            get_output_lib_dir = "../../lib/win32"
        End If
    End Function

    Private Function parse_by_type(type_mode)
        Dim node, nodeOutDir, nodeIntDir, nodeAdditionalIncludeDirectories, nodePreprocessorDefinitions, nodeOutputFile, nodeAdditionalLibraryDirectories
        If doc_.parseError.errorCode = 0 Then
            '------------------------------------------------
            ' Release              


            Set node = root_.selectSingleNode(get_xpath_by_type("//PropertyGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""][not(@Label)]", type_mode))
            If Not (node is Nothing) Then              
                
                    
                Set nodeOutDir = root_.selectSingleNode(get_xpath_by_type("//PropertyGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""][not(@Label)]/OutDir", type_mode))
                If Not (nodeOutDir is Nothing) Then
                    nodeOutDir.text = "..\..\tmp\$(ConfigurationName)\$(ProjectName)\"
                Else
                    Set nodeOutDir = doc_.createElement("OutDir")			
                    nodeOutDir.text = "..\..\tmp\$(ConfigurationName)\$(ProjectName)\"			
                    node.appendChild(nodeOutDir)			
                End If 
                
                Set nodeIntDir = root_.selectSingleNode(get_xpath_by_type("//PropertyGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""][not(@Label)]/IntDir", type_mode))
                If Not (nodeIntDir is Nothing) Then
                    nodeIntDir.text = "..\..\tmp\$(ConfigurationName)\$(ProjectName)\"
                Else
                    Set nodeIntDir = doc_.createElement("IntDir")			
                    nodeIntDir.text = "..\..\tmp\$(ConfigurationName)\$(ProjectName)\"			
                    node.appendChild(nodeIntDir)		
                End If 


            End If 
            
            
            Set node = root_.selectSingleNode(get_xpath_by_type("//ItemDefinitionGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""]//ClCompile", type_mode))
            If Not (node is Nothing) Then     


                Set nodeAdditionalIncludeDirectories = root_.selectSingleNode(get_xpath_by_type("//ItemDefinitionGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""]//ClCompile/AdditionalIncludeDirectories", type_mode))
                If Not (nodeAdditionalIncludeDirectories is Nothing) Then
                    additional_include_directories_ = Replace(nodeAdditionalIncludeDirectories.text, "%(AdditionalIncludeDirectories)", "")                    
                Else
                    Set nodeAdditionalIncludeDirectories = doc_.createElement("AdditionalIncludeDirectories")			
                    nodeAdditionalIncludeDirectories.text = "../../publish;../../include;../../3rdparty;../../3rdparty/msvc_compat;../../3rdparty/wtl;../../3rdparty/htmlayout"			
                    node.appendChild(nodeAdditionalIncludeDirectories)		
                    additional_include_directories_ = Replace(nodeAdditionalIncludeDirectories.text, "%(AdditionalIncludeDirectories)", "")     
                End If 		

                Set nodePreprocessorDefinitions = root_.selectSingleNode(get_xpath_by_type("//ItemDefinitionGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""]//ClCompile/PreprocessorDefinitions", type_mode))
                If Not (nodePreprocessorDefinitions is Nothing) Then
                    nodePreprocessorDefinitions.text = 	AddPreDefinitions(nodePreprocessorDefinitions.text, "_CRT_SECURE_NO_DEPRECATE")
                    preprocessor_definitions_ = Replace(nodePreprocessorDefinitions.text, "%(PreprocessorDefinitions)", "")
                End If 
            End If 
            
            Set node = root_.selectSingleNode(get_xpath_by_type("//ItemDefinitionGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""]/Link", type_mode))
            If Not (node is Nothing) Then
                Set nodeOutputFile = root_.selectSingleNode(get_xpath_by_type("//ItemDefinitionGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""]/Link/OutputFile", type_mode))
                If Not (nodeOutputFile is Nothing) Then
                    nodeOutputFile.text = get_output_file(type_mode)
                Else
                    Set nodeOutputFile = doc_.createElement("OutputFile")			
                    nodeOutputFile.text = get_output_file(type_mode)		
                    node.appendChild(nodeOutputFile)			
                End If 		


                Set nodeAdditionalLibraryDirectories = root_.selectSingleNode(get_xpath_by_type("//ItemDefinitionGroup[@Condition=""'$(Configuration)|$(Platform)'=='Release|Win32'""]/Link/AdditionalLibraryDirectories", type_mode))
                If Not (nodeAdditionalLibraryDirectories is Nothing) Then
                    nodeAdditionalLibraryDirectories.text = get_output_lib_dir(type_mode)
                Else
                    Set nodeAdditionalLibraryDirectories = doc_.createElement("AdditionalLibraryDirectories")			
                    nodeAdditionalLibraryDirectories.text = get_output_lib_dir(type_mode)	
                    node.appendChild(nodeAdditionalLibraryDirectories)			
                End If 	
            End If 
            
            '------------------------------------------------
            ' 资源文件
            Set resfiles = root_.selectNodes("//ItemGroup//None")
            If Not (resfiles is Nothing) Then		
                For I = 0 To resfiles.length-1
                    WScript.Echo resfiles(I).getAttribute("Include") 	
                Next	
            End If 
            
            Set resfiles2 = root_.selectNodes("//ItemGroup//ResourceCompile")
            If Not (resfiles2 is Nothing) Then		
                For I = 0 To resfiles2.length-1
                    WScript.Echo resfiles2(I).getAttribute("Include") 	
                Next	
            End If 

            '------------------------------------------------
            ' 头文件
            Set hfiles = root_.selectNodes("//ItemGroup//ClInclude")
            If Not (hfiles is Nothing) Then		
                For I = 0 To hfiles.length-1
                    WScript.Echo hfiles(I).getAttribute("Include") 	
                Next	
            End If 
            
            '------------------------------------------------
            ' C++文件
            Set cfiles = root_.selectNodes("//ItemGroup//ClCompile")
            If Not (cfiles is Nothing) Then		
                For I = 0 To cfiles.length-1
                    WScript.Echo cfiles(I).getAttribute("Include") 	
                Next	
            End If

        End If
    End Function

    Public Function parse()
        Call parse_by_type("Debug")
        Call parse_by_type("Release")
        doc_.save(filename_)        
    End Function

    Public Function export_vc2005_doc()
        Dim include_node, output_node, res_node, midl_node
        Dim doc, root, prjroot, croot, hroot, resroot, nodefile, I
        Set doc = CreateObject("MSXML2.DOMDocument")
        doc.async = False
        doc.load("vc2005_template.vcproj")
        If doc.parseError.errorCode = 0 Then
            Set root = doc.documentElement		

            set prjroot = root.selectSingleNode("//VisualStudioProject")
            If not (prjroot is Nothing) Then
                prjroot.setAttribute("Name") = get_root_namespace()
                prjroot.setAttribute("ProjectGUID") = get_project_guid()
                prjroot.setAttribute("RootNamespace") = get_root_namespace()
                prjroot.setAttribute("Keyword") = get_keyword()
            End If
            
            ' 添加C++文件
            Set croot = root.selectSingleNode("//Files//Filter[@Name=""源文件""]")
            If Not (croot is Nothing) Then
                For I = 0 To cfiles.length-1
                    Set nodefile = doc.createElement("File")			
                    nodefile.setAttribute("RelativePath") = cfiles(I).getAttribute("Include") 				

                    If cfiles(I).getAttribute("Include") = "stdafx.cpp" Then
                        Set nodeFileConfiguration = doc.createElement("FileConfiguration")
                        nodeFileConfiguration.setAttribute("Name") = "Release|Win32"

                        Set nodeTool = doc.createElement("Tool")
                        nodeTool.setAttribute("Name") = "VCCLCompilerTool"
                        nodeTool.setAttribute("UsePrecompiledHeader") = "1"

                        nodeFileConfiguration.appendChild(nodeTool)
                        nodeFile.appendChild(nodeFileConfiguration)
                    End If

                    croot.appendChild(nodefile)							
                Next			
            End If 
            
            ' 添加头文件
            Set hroot = root.selectSingleNode("//Files//Filter[@Name=""头文件""]")
            If Not (hroot is Nothing) Then
                For I = 0 To hfiles.length-1
                    Set nodefile = doc.createElement("File")			
                    nodefile.setAttribute("RelativePath") = hfiles(I).getAttribute("Include")								 				
                    hroot.appendChild(nodefile)				
                                    
                Next			
            End If 
            
            ' 添加资源文件
            Set resroot = root.selectSingleNode("//Files//Filter[@Name=""资源文件""]")
            If Not (resroot is Nothing) Then
                For I = 0 To resfiles.length-1
                    Set nodefile = doc.createElement("File")			
                    nodefile.setAttribute("RelativePath") = resfiles(I).getAttribute("Include") 	
                    
                    If resfiles(I).getAttribute("Include") = "ReadMe.txt" Then
                        root.selectSingleNode("//Files").appendChild(nodefile)	
                    Else					 				
                        resroot.appendChild(nodefile)				
                    End If							
                Next			

                For I = 0 To resfiles2.length-1
                    Set nodefile = doc.createElement("File")			
                    nodefile.setAttribute("RelativePath") = resfiles2(I).getAttribute("Include") 				
                    resroot.appendChild(nodefile)				
                Next
            End If 

            Set node = root.selectSingleNode("//Configuration[@Name=""Release|Win32""]")
            If Not (node is Nothing) Then
                WScript.Echo node.getAttribute("Name")
                node.setAttribute("OutputDirectory") = "../../tmp/$(ConfigurationName)/$(ProjectName)"
                node.setAttribute("IntermediateDirectory") = "../../tmp/$(ConfigurationName)/$(ProjectName)"
                node.setAttribute("BuildLogFile") = "../../tmp/$(ConfigurationName)/$(ProjectName)/BuildLog.htm"
                node.setAttribute("ATLMinimizesCRunTimeLibraryUsage") = "false"
                If get_config_type("Release") = "StaticLibrary" Then
                    node.setAttribute("ConfigurationType") = "4"
                ElseIf get_config_type("Release") = "DynamicLibrary" Then
                    node.setAttribute("ConfigurationType") = "2"
                ElseIf get_config_type("Release") = "Application" Then
                    node.setAttribute("ConfigurationType") = "1"
                End If
                config_type = CStr(node.getAttribute("ConfigurationType"))
            End If

            Set include_node = root.selectSingleNode("//Configuration[@Name=""Release|Win32""]//Tool[@Name=""VCCLCompilerTool""]")
            If Not (include_node is Nothing) Then
                include_node.setAttribute("PrecompiledHeaderFile") = "$(IntDir)\$(TargetName).pch"
                include_node.setAttribute("AssemblerListingLocation") = "$(IntDir)\"
                include_node.setAttribute("ObjectFile") = "$(IntDir)\"
                include_node.setAttribute("ProgramDataBaseFileName") = "$(IntDir)\vc80.pdb"
                include_node.setAttribute("AdditionalIncludeDirectories") = additional_include_directories_ & ";../../publish;../../include;../../3rdparty;../../3rdparty/msvc_compat;../../3rdparty/wtl;../../3rdparty/htmlayout"	
                include_node.setAttribute("UsePrecompiledHeader") = get_precompiled_header("Release")
                include_node.setAttribute("PreprocessorDefinitions") = preprocessor_definitions_
                
            End If

            Set output_node = root.selectSingleNode("//Configuration[@Name=""Release|Win32""]//Tool[@Name=""VCLinkerTool""]")
            If Not (output_node is Nothing) Then
                If config_type = "1" Then
                    output_node.setAttribute("OutputFile") = "../../product/win32/$(ProjectName).exe"
                ElseIf  config_type = "2" Then
                    output_node.setAttribute("OutputFile") = "../../product/win32/$(ProjectName).dll"	
                ElseIf  config_type = "4" Then
                    output_node.setAttribute("OutputFile") = "../../lib/win32/$(ProjectName).lib"
                End If
                
                output_node.setAttribute("AdditionalLibraryDirectories") = "../../3rdparty/gl;../../lib/win32;../../3rdparty/Python27/libs"	
                output_node.setAttribute("GenerateDebugInformation") = "true"	
                output_node.setAttribute("ProgramDatabaseFile") = "../../product/win32/dbginfo/$(ProjectName).pdb"	
            End If 	
            
            Set res_node = root.selectSingleNode("//Configuration[@Name=""Release|Win32""]//Tool[@Name=""VCResourceCompilerTool""]")
            If Not (res_node is Nothing) Then
                res_node.setAttribute("AdditionalIncludeDirectories") = "../../3rdparty/wtl;$(IntDir)"			
            End If 
            
            Set midl_node = root.selectSingleNode("//Configuration[@Name=""Release|Win32""]//Tool[@Name=""VCMIDLTool""]")
            If Not (midl_node is Nothing) Then
                midl_node.setAttribute("TypeLibraryName") = "$(IntDir)\$(TargetName).tlb"			
            End If

            If config_type = "4" Then
                Set tool_node = root.selectSingleNode("//Configuration[@Name=""Release|Win32""]//Tool[@Name=""VCLibrarianTool""]")
                If tool_node is Nothing Then
                    Set tool_node = doc.createElement("Tool")			
                    tool_node.setAttribute("Name") = "VCLibrarianTool" 
                    tool_node.setAttribute("OutputFile") = "../../lib/win32/$(ProjectName)_md.lib" 
                    node.appendChild(tool_node)                    
                End If 
                     	
            End If



        End If 
        
        doc.save(GetFilePath(filename_) & "\" & GetBaseName(filename_) & "-vc8.vcproj")

        Call ReplaceFileContent(filename_, " xmlns=""""", "", 1) 
    End Function

    Private Function AddPreDefinitions(sourcestr, substr)
        Dim temparr, I	
        temparr = Split(sourcestr, ";")
        For I = 0 To UBound(temparr)
            If temparr(I) = substr Then		
                AddPreDefinitions = sourcestr
                Exit Function
            End If
        Next

        AddPreDefinitions = substr & ";" & sourcestr
    End Function
End Class


Class LibLoader    
    Private lib_dir_
    
    Private Sub Class_Initialize()
        Dim objShell
        lib_dir_ = left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))        
        Set objShell = wscript.createObject("wscript.shell")
        objShell.CurrentDirectory = lib_dir_
    End Sub
    
    Private Sub Class_Terminate()        
    End Sub

    Public Property Get Path
        Path = lib_dir_
    End Property
    
    Public Property Let Path(value)
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        lib_dir_ = fso.GetAbsolutePathName(value)
        Set fso = Nothing
    End Property
    
    Public Function Import(ByVal filename) 
        Dim fso, sh, file, code, dir, basename

        ' Create my own objects, so the function is self-contained and can be called
        ' before anything else in the script.
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")

        filename = Trim(sh.ExpandEnvironmentStrings(filename))
        If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
            ' filename is not absolute
            If Not fso.FileExists(fso.GetAbsolutePathName(filename)) Then
                If fso.FileExists(fso.BuildPath(lib_dir_, filename)) Then
                    filename = fso.BuildPath(lib_dir_, filename)                    
                End If                
            End If
            filename = fso.GetAbsolutePathName(filename)
        End If

        'WScript.Echo filename

        On Error Resume Next
        Set file = fso.GetFile(filename)
        basename = fso.GetBaseName(file)
        ExecuteGlobal "Const " & basename & "_vbs_loading = 1"
        If Err = 0 Then
            On Error Resume Next
            Set file = fso.OpenTextFile(filename, 1, False)
            If Err Then
                WScript.Echo "Cannot import '" & filename & "': " & Err.Description & " (0x" & Hex(Err.Number) & ")"
                WScript.Quit 1
            End If
            On Error Goto 0
            code = file.ReadAll
            file.Close
            ExecuteGlobal(code)        
        End If
        Set sh = Nothing
        Set fso = Nothing
    End Function    
End Class