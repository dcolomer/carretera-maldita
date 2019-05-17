Attribute VB_Name = "DError"
Option Explicit

Enum DirectXMethod
    Direct3D_IM = 0
    Direct3D_RM = 1
    DirectDraw = 2
    DirectInput = 3
    DirectMusic = 4
    DirectPlay = 5
    DirectSound = 6
End Enum

' Se pasa Err.Number, el tipo de objeto DirectX y el
' lugar donde se ha producido el error
Public Sub HandleAnyErrors(ByVal ErrorNum As Long, ByVal DirectXType As DirectXMethod, ByVal ProcedimientoConError As String)

  ' Si ha ocurrido un error
  If ErrorNum <> 0 Then

    '-------------------------------------------------
    
    If Not (perf Is Nothing) Then perf.CloseDown
    If Not (DS Is Nothing) Then Set DS = Nothing
    
    DD.RestoreDisplayMode
    DD.SetCooperativeLevel frmPrincipal.hWnd, DDSCL_NORMAL
    
    Set DD = Nothing
    Set DX = Nothing
       
    
    ' Si el error es de lectura de algun fichero de disco
    If ProcedimientoConError = "LeerMapaDisco" Or ProcedimientoConError = "GuardarRecord" Then
      
      MsgBox "Se ha producido un error al acceder a un fichero de disco.", vbExclamation, "Error"

    ' Si el error es DicrectX
    Else
    
      '-------------------------------------------------
      ' Hacer que error sea inteligible para el usuario.
      Dim ErrID As String, ErrDesc As String, DXInstance As String
    
      GetErrorDesc ErrorNum, ErrID, ErrDesc, DirectXType
    
      Select Case DirectXType
        Case Direct3D_IM: DXInstance = "Direct3D_IM"
        Case Direct3D_RM: DXInstance = "Direct3D_RM"
        Case DirectDraw: DXInstance = "DirectDraw"
        Case DirectInput: DXInstance = "DirectInput"
        Case DirectMusic: DXInstance = "DirectMusic"
        Case DirectPlay: DXInstance = "DirectPlay"
        Case DirectSound: DXInstance = "DirectSound"
      End Select
    
      ' Informar al usuario del error
      MsgBox "Ha ocurrido un error en " & DXInstance & vbCrLf & vbCrLf & _
             "Número de error: " & ErrorNum & vbCrLf & "Id. de error: " & ErrID & vbCrLf & _
             "Descripción del error: " & vbCrLf & _
             "Ha ocurrido en: " & ProcedimientoConError, vbCritical, "Error"
    End If
  End If

End Sub

Public Sub GetErrorDesc(ByVal TempErr As Long, ByRef ErrorID As String, ByRef Description As String, ByVal DXmod As DirectXMethod)
    
    Select Case DXmod
        Case Direct3D_IM
            Select Case TempErr
                Case D3D_OK
                    ErrorID = "D3D_OK"
                    Description = "No error occurred."
                Case D3DERR_BADMAJORVERSION
                    ErrorID = "D3DERR_BADMAJORVERSION"
                    Description = "The service you requested is unavailable in this major version of DirectX. (A " + Chr$(34) + "major version" + Chr$(34) + " denotes a primary release, such as DirectX 6.0.)"
                Case D3DERR_BADMINORVERSION
                    ErrorID = "D3DERR_BADMINORVERSION"
                    Description = "The service you requested is available in this major version of DirectX, but not in this minor version. Get the latest version of the component runtime from Microsoft. (A " + Chr$(34) + "minor version" + Chr$(34) + " denotes a secondary release, such as DirectX 6.1.)"
                Case D3DERR_COLORKEYATTACHED
                    ErrorID = "D3DERR_COLORKEYATTACHED"
                    Description = "The application attempted to create a texture with a surface that uses a color key for transparency."
                Case D3DERR_CONFLICTINGTEXTUREFILTER
                    ErrorID = "D3DERR_CONFLICTINGTEXTUREFILTER"
                    Description = "The current texture filters cannot be used together."
                Case D3DERR_CONFLICTINGTEXTUREPALETTE
                    ErrorID = "D3DERR_CONFLICTINGTEXTUREPALETTE"
                    Description = "The current textures cannot be used simultaneously. This generally occurs when a multitexture device requires that all palettized textures simultaneously enabled also share the same palette."
                Case D3DERR_CONFLICTINGRENDERSTATE
                    ErrorID = "D3DERR_CONFLICTINGRENDERSTATE"
                    Description = "The currently set render states cannot be used together."
                Case D3DERR_DEVICEAGGREGATED
                    ErrorID = "D3DERR_DEVICEAGGREGATED"
                    Description = "The Direct3DDevice7.SetRenderTarget method was called on a device that was retrieved from the render target surface."
                Case D3DERR_INITFAILED
                    ErrorID = "D3DERR_INITFAILED"
                    Description = "A rendering device could not be created because the new device could not be initialized."
                Case D3DERR_INBEGIN
                    ErrorID = "D3DERR_INBEGIN"
                    Description = "The requested operation cannot be completed while scene rendering is taking place. Try again after the scene is completed and the Direct3DDevice7.EndScene method (or equivalent method) is called."
                Case D3DERR_INBEGINSTATEBLOCK
                    ErrorID = "D3DERR_INBEGINSTATEBLOCK"
                    Description = "The operation cannot be completed while recording states for a state block. Complete recording by calling the Direct3DDevice7.EndStateBlock method and try again."
                'Case D3DERR_INOVERLAYSTATEBLOCK
                 '   ErrorID = "D3DERR_INOVERLAYSTATEBLOCK"
                 '   Description = "The operation cannot be completed while overlaying a state block. Remove the state block overlay and try again."
                Case D3DERR_INVALID_DEVICE
                    ErrorID = "D3DERR_INVALID_DEVICE"
                    Description = "The requested device type is not valid."
                Case D3DERR_INVALIDCURRENTVIEWPORT
                    ErrorID = "D3DERR_INVALIDCURRENTVIEWPORT"
                    Description = "The currently selected viewport is not valid."
                Case D3DERR_INVALIDMATRIX
                    ErrorID = "D3DERR_INVALIDMATRIX"
                    Description = "The requested operation could not be completed because the combination of the currently set world, view, and projection matrices is invalid (the determinant of the combined matrix is zero)."
                Case D3DERR_INVALIDPALETTE
                    ErrorID = "D3DERR_INVALIDPALETTE"
                    Description = "The palette associated with a surface is invalid."
                Case D3DERR_INVALIDPRIMITIVETYPE
                    ErrorID = "D3DERR_INVALIDPRIMITIVETYPE"
                    Description = "The primitive type specified by the application is invalid."
                Case D3DERR_INVALIDRAMPTEXTURE
                    ErrorID = "D3DERR_INVALIDRAMPTEXTURE"
                    Description = "Ramp mode is being used and the texture handle in the current material does not match the current texture handle that is set as a render state."
                Case D3DERR_INVALIDSTATEBLOCK
                    ErrorID = "D3DERR_INVALIDSTATEBLOCK"
                    Description = "The state block handle is invalid."
                Case D3DERR_INVALIDVERTEXFORMAT
                    ErrorID = "D3DERR_INVALIDVERTEXFORMAT"
                    Description = "The combination of flexible vertex format flags specified by the application is not valid."
                Case D3DERR_INVALIDVERTEXTYPE
                    ErrorID = "D3DERR_INVALIDVERTEXTYPE"
                    Description = "The vertex type specified by the application is invalid."
                Case D3DERR_LIGHT_SET_FAILED
                    ErrorID = "D3DERR_LIGHT_SET_FAILED"
                    Description = "The attempt to set lighting parameters for a light object failed."
                Case D3DERR_LIGHTHASVIEWPORT
                    ErrorID = "D3DERR_LIGHTHASVIEWPORT"
                    Description = "The requested operation failed because the light object is associated with another viewport."
                Case D3DERR_LIGHTNOTINTHISVIEWPORT
                    ErrorID = "D3DERR_LIGHTNOTINTHISVIEWPORT"
                    Description = "The requested operation failed because the light object has not been associated with this viewport."
                Case D3DERR_MATERIAL_CREATE_FAILED
                    ErrorID = "D3DERR_MATERIAL_CREATE_FAILED"
                    Description = "The material could not be created. This typically occurs when no memory is available to allocate for the material."
                Case D3DERR_MATERIAL_DESTROY_FAILED
                    ErrorID = "D3DERR_MATERIAL_DESTROY_FAILED"
                    Description = "The memory for the material could not be deallocated."
                Case D3DERR_MATERIAL_GETDATA_FAILED
                    ErrorID = "D3DERR_MATERIAL_GETDATA_FAILED"
                    Description = "The material parameters could not be retrieved."
                Case D3DERR_MATERIAL_SETDATA_FAILED
                    ErrorID = "D3DERR_MATERIAL_SETDATA_FAILED"
                    Description = "The material parameters could not be set."
                Case D3DERR_MATRIX_CREATE_FAILED
                    ErrorID = "D3DERR_MATRIX_CREATE_FAILED"
                    Description = "The matrix could not be created. This can occur when no memory is available to allocate for the matrix."
                Case D3DERR_MATRIX_DESTROY_FAILED
                    ErrorID = "D3DERR_MATRIX_DESTROY_FAILED"
                    Description = "The memory for the matrix could not be deallocated."
                Case D3DERR_MATRIX_GETDATA_FAILED
                    ErrorID = "D3DERR_MATRIX_GETDATA_FAILED"
                    Description = "The matrix data could not be retrieved. This can occur when the matrix was not created by the current device."
                Case D3DERR_MATRIX_SETDATA_FAILED
                    ErrorID = "D3DERR_MATRIX_SETDATA_FAILED"
                    Description = "The matrix data could not be set. This can occur when the matrix was not created by the current device."
                Case D3DERR_NOCURRENTVIEWPORT
                    ErrorID = "D3DERR_NOCURRENTVIEWPORT"
                    Description = "The viewport parameters could not be retrieved because none have been set."
                Case D3DERR_NOTINBEGIN
                    ErrorID = "D3DERR_NOTINBEGIN"
                    Description = "The requested rendering operation could not be completed because scene rendering has not begun. Call Direct3DDevice7.BeginScene to begin rendering then try again."
                Case D3DERR_NOTINBEGINSTATEBLOCK
                    ErrorID = "D3DERR_NOTINBEGINSTATEBLOCK"
                    Description = "The requested operation could not be completed because it is only valid while recording a state block. Call the Direct3DDevice7.BeginStateBlock method and try again."
                Case D3DERR_NOVIEWPORTS
                    ErrorID = "D3DERR_NOVIEWPORTS"
                    Description = "The requested operation failed because the device currently has no viewports associated with it."
                Case D3DERR_SCENE_BEGIN_FAILED
                    ErrorID = "D3DERR_SCENE_BEGIN_FAILED"
                    Description = "Scene rendering could not begin."
                Case D3DERR_SCENE_END_FAILED
                    ErrorID = "D3DERR_SCENE_END_FAILED"
                    Description = "Scene rendering could not be completed."
                Case D3DERR_SCENE_IN_SCENE
                    ErrorID = "D3DERR_SCENE_IN_SCENE"
                    Description = "Scene rendering could not begin because a previous scene was not completed by a call to the Direct3DDevice7.EndScene method."
                Case D3DERR_SCENE_NOT_IN_SCENE
                    ErrorID = "D3DERR_SCENE_NOT_IN_SCENE"
                    Description = "Scene rendering could not be completed because a scene was not started by a previous call to the Direct3DDevice7.BeginScene method."
                Case D3DERR_SETVIEWPORTDATA_FAILED
                    ErrorID = "D3DERR_SETVIEWPORTDATA_FAILED"
                    Description = "The viewport parameters could not be set."
                Case D3DERR_STENCILBUFFER_NOTPRESENT
                    ErrorID = "D3DERR_STENCILBUFFER_NOTPRESENT"
                    Description = "The requested stencil buffer operation could not be completed because there is no stencil buffer attached to the render target surface."
                Case D3DERR_SURFACENOTINVIDMEM
                    ErrorID = "D3DERR_SURFACENOTINVIDMEM"
                    Description = "The device could not be created because the render target surface is not located in video-memory. (Hardware-accelerated devices require video-memory render target surfaces.)"
                Case D3DERR_TEXTURE_BADSIZE
                    ErrorID = "D3DERR_TEXTURE_BADSIZE"
                    Description = "The dimensions of a current texture are invalid. This can occur when an application attempts to use a texture that has non-power-of-two dimensions with a device that requires them."
                Case D3DERR_TEXTURE_CREATE_FAILED
                    ErrorID = "D3DERR_TEXTURE_CREATE_FAILED"
                    Description = "The texture handle for the texture could not be retrieved from the driver."
                Case D3DERR_TEXTURE_DESTROY_FAILED
                    ErrorID = "D3DERR_TEXTURE_DESTROY_FAILED"
                    Description = "The device was unable to deallocate the texture memory."
                Case D3DERR_TEXTURE_GETSURF_FAILED
                    ErrorID = "D3DERR_TEXTURE_GETSURF_FAILED"
                    Description = "The DirectDraw surface used to create the texture could not be retrieved."
                Case D3DERR_TEXTURE_LOAD_FAILED
                    ErrorID = "D3DERR_TEXTURE_LOAD_FAILED"
                    Description = "The texture could not be loaded."
                Case D3DERR_TEXTURE_LOCK_FAILED
                    ErrorID = "D3DERR_TEXTURE_LOCK_FAILED"
                    Description = "The texture could not be locked."
                Case D3DERR_TEXTURE_LOCKED
                    ErrorID = "D3DERR_TEXTURE_LOCKED"
                    Description = "The requested operation could not be completed because the texture surface is currently locked."
                Case D3DERR_TEXTURE_NO_SUPPORT
                    ErrorID = "D3DERR_TEXTURE_NO_SUPPORT"
                    Description = "The device does not support texture mapping."
                Case D3DERR_TEXTURE_NOT_LOCKED
                    ErrorID = "D3DERR_TEXTURE_NOT_LOCKED"
                    Description = "The requested operation could not be completed because the texture surface is not locked."
                Case D3DERR_TEXTURE_SWAP_FAILED
                    ErrorID = "D3DERR_TEXTURE_SWAP_FAILED"
                    Description = "The texture handles could not be swapped."
                Case D3DERR_TEXTURE_UNLOCK_FAILED
                    ErrorID = "D3DERR_TEXTURE_UNLOCK_FAILED"
                    Description = "The texture surface could not be unlocked."
                Case D3DERR_TOOMANYOPERATIONS
                    ErrorID = "D3DERR_TOOMANYOPERATIONS"
                    Description = "The application is requesting more texture filtering operations than the device supports."
                Case D3DERR_TOOMANYPRIMITIVES
                    ErrorID = "D3DERR_TOOMANYPRIMITIVES"
                    Description = "The device is unable to render the provided quantity of primitives in a single pass."
                Case D3DERR_UNSUPPORTEDALPHAARG
                    ErrorID = "D3DERR_UNSUPPORTEDALPHAARG"
                    Description = "The device does not support one of the specified texture blending arguments for the alpha channel."
                Case D3DERR_UNSUPPORTEDALPHAOPERATION
                    ErrorID = "D3DERR_UNSUPPORTEDALPHAOPERATION"
                    Description = "The device does not support one of the specified texture blending operations for the alpha channel."
                Case D3DERR_UNSUPPORTEDCOLORARG
                    ErrorID = "D3DERR_UNSUPPORTEDCOLORARG"
                    Description = "The device does not support the one of the specified texture blending arguments for color values."
                Case D3DERR_UNSUPPORTEDCOLOROPERATION
                    ErrorID = "D3DERR_UNSUPPORTEDCOLOROPERATION"
                    Description = "The device does not support the one of the specified texture blending operations for color values."
                Case D3DERR_UNSUPPORTEDFACTORVALUE
                    ErrorID = "D3DERR_UNSUPPORTEDFACTORVALUE"
                    Description = "The specified texture factor value is not supported by the device."
                Case D3DERR_UNSUPPORTEDTEXTUREFILTER
                    ErrorID = "D3DERR_UNSUPPORTEDTEXTUREFILTER"
                    Description = "The specified texture filter is not supported by the device."
                Case D3DERR_VBUF_CREATE_FAILED
                    ErrorID = "D3DERR_VBUF_CREATE_FAILED"
                    Description = "The vertex buffer could not be created. This can happen when there is insufficient memory to allocate a vertex buffer."
                Case D3DERR_VERTEXBUFFERLOCKED
                    ErrorID = "D3DERR_VERTEXBUFFERLOCKED"
                    Description = "The requested operation could not be completed because the vertex buffer is locked."
                Case D3DERR_VERTEXBUFFEROPTIMIZED
                    ErrorID = "D3DERR_VERTEXBUFFEROPTIMIZED"
                    Description = "The requested operation could not be completed because the vertex buffer is optimized. (The contents of optimized vertex buffers are driver specific, and considered private.)"
                Case D3DERR_VERTEXBUFFERUNLOCKFAILED
                    ErrorID = "D3DERR_VERTEXBUFFERUNLOCKFAILED"
                    Description = "The vertex buffer could not be unlocked because the vertex buffer memory was overrun. Make sure that your application does not write beyond the size of the vertex buffer."
                Case D3DERR_VIEWPORTDATANOTSET
                    ErrorID = "D3DERR_VIEWPORTDATANOTSET"
                    Description = "The requested operation could not be completed because viewport parameters have not yet been set. Set the viewport parameters by calling Direct3DDevice7.SetViewport method and try again."
                Case D3DERR_VIEWPORTHASNODEVICE
                    ErrorID = "D3DERR_VIEWPORTHASNODEVICE"
                    Description = "The requested operation could not be completed because the viewport has not yet been associated with a device."
                Case D3DERR_WRONGTEXTUREFORMAT
                    ErrorID = "D3DERR_WRONGTEXTUREFORMAT"
                    Description = "The pixel format of the texture surface is not valid."
                Case D3DERR_ZBUFF_NEEDS_SYSTEMMEMORY
                    ErrorID = "D3DERR_ZBUFF_NEEDS_SYSTEMMEMORY"
                    Description = "The requested operation could not be completed because the specified device requires system-memory depth-buffer surfaces. (Software rendering devices require system-memory depth buffers.)"
                Case D3DERR_ZBUFF_NEEDS_VIDEOMEMORY
                    ErrorID = "D3DERR_ZBUFF_NEEDS_VIDEOMEMORY"
                    Description = "The requested operation could not be completed because the specified device requires video-memory depth-buffer surfaces. (Hardware-accelerated devices require video-memory depth buffers.)"
                Case D3DERR_ZBUFFER_NOTPRESENT
                    ErrorID = "D3DERR_ZBUFFER_NOTPRESENT"
                    Description = "The requested operation could not be completed because the render target surface does not have an attached depth buffer."
            End Select
        Case Direct3D_RM
            Select Case TempErr
                Case D3DRM_OK
                    ErrorID = "D3DRM_OK"
                    Description = "No error. Equivalent to DD_OK."
                Case D3DRMERR_BADALLOC
                    ErrorID = "D3DRMERR_BADALLOC"
                    Description = "Out of memory."
                Case D3DRMERR_BADDEVICE
                    ErrorID = "D3DRMERR_BADDEVICE"
                    Description = "Device is not compatible with renderer."
                Case D3DRMERR_BADFILE
                    ErrorID = "D3DRMERR_BADFILE"
                    Description = "Data file is corrupt."
                Case D3DRMERR_BADMAJORVERSION
                    ErrorID = "D3DRMERR_BADMAJORVERSION"
                    Description = "Bad dynamic-link library (DLL) major version."
                Case D3DRMERR_BADMINORVERSION
                    ErrorID = "D3DRMERR_BADMINORVERSION"
                    Description = "Bad DLL minor version."
                Case D3DRMERR_BADOBJECT
                    ErrorID = "D3DRMERR_BADOBJECT"
                    Description = "Object expected in argument."
                Case D3DRMERR_BADPMDATA
                    ErrorID = "D3DRMERR_BADPMDATA"
                    Description = "Data in the .x file is corrupted. The conversion to a progressive mesh succeeded but produced an invalid progressive mesh in the .x file."
                Case D3DRMERR_BADTYPE
                    ErrorID = "D3DRMERR_BADTYPE"
                    Description = "Bad argument type passed."
                Case D3DRMERR_BADVALUE
                    ErrorID = "D3DRMERR_BADVALUE"
                    Description = "Bad argument value passed."
                Case D3DRMERR_BOXNOTSET
                    ErrorID = "D3DRMERR_BOXNOTSET"
                    Description = "An attempt was made to access a bounding box (for example, with Direct3DRMFrame3.GetBox) when no bounding box was set on the frame."
                Case D3DRMERR_CLIENTNOTREGISTERED
                    ErrorID = "D3DRMERR_CLIENTNOTREGISTERED"
                    Description = "Client has not been registered."
                Case D3DRMERR_CONNECTIONLOST
                    ErrorID = "D3DRMERR_CONNECTIONLOST"
                    Description = "Data connection was lost during a load, clone, or duplicate."
                Case D3DRMERR_ELEMENTINUSE
                    ErrorID = "D3DRMERR_ELEMENTINUSE"
                    Description = "Element can't be modified or deleted while in use. To empty a submesh, call Direct3DRMMeshBuilder3.Empty against its parent."
                'Case D3DRMERR_ENTRYINUSE
                '    ErrorID = "D3DRMERR_ENTRYINUSE"
                '    Description = "Vertex or normal entries are in use by a face and cannot be deleted."
                Case D3DRMERR_FACEUSED
                    ErrorID = "D3DRMERR_FACEUSED"
                    Description = "Face already used in a mesh."
                Case D3DRMERR_FILENOTFOUND
                    ErrorID = "D3DRMERR_FILENOTFOUND"
                    Description = "File not found in the specified location."
                'Case D3DRMERR_INCOMPATIBLEKEY
                '    ErrorID = "D3DRMERR_INCOMPATIBLEKEY"
                '    Description = "Specified animation key is incompatible. The key cannot be modified."
                Case D3DRMERR_INVALIDLIBRARY
                    ErrorID = "D3DRMERR_INVALIDLIBRARY"
                    Description = "Specified library is invalid."
                'Case D3DRMERR_INVALIDOBJECT
                '    ErrorID = "D3DRMERR_INVALIDOBJECT"
                '    Description = "Method received an object that is invalid."
                'Case D3DRMERR_INVALIDPARAMS
                '    ErrorID = "D3DRMERR_INVALIDPARAMS"
                '    Description = "At least one of the parameters passed to the method is invalid."
                Case D3DRMERR_LIBRARYNOTFOUND
                    ErrorID = "D3DRMERR_LIBRARYNOTFOUND"
                    Description = "Specified library not found."
                Case D3DRMERR_LOADABORTED
                    ErrorID = "D3DRMERR_LOADABORTED"
                    Description = "Load aborted by user."
                Case D3DRMERR_NOSUCHKEY
                    ErrorID = "D3DRMERR_NOSUCHKEY"
                    Description = "Specified animation key does not exist."
                Case D3DRMERR_NOTCREATEDFROMDDS
                    ErrorID = "D3DRMERR_NOTCREATEDFROMDDS"
                    Description = "Specified texture was not created from a Microsoft DirectDraw® surface."
                Case D3DRMERR_NOTDONEYET
                    ErrorID = "D3DRMERR_NOTDONEYET"
                    Description = "Error flag not implemented."
                Case D3DRMERR_NOTENOUGHDATA
                    ErrorID = "D3DRMERR_NOTENOUGHDATA"
                    Description = "Not enough data has been loaded to perform the requested operation."
                Case D3DRMERR_NOTFOUND
                    ErrorID = "D3DRMERR_NOTFOUND"
                    Description = "Object not found in specified place."
                'Case D3DRMERR_OUTOFRANGE
                '    ErrorID = "D3DRMERR_OUTOFRANGE"
                '    Description = "Specified value is out of range."
                Case D3DRMERR_PENDING
                    ErrorID = "D3DRMERR_PENDING"
                    Description = "Data required to supply the requested information has not finished loading."
                Case D3DRMERR_REQUESTTOOLARGE
                    ErrorID = "D3DRMERR_REQUESTTOOLARGE"
                    Description = "Attempt was made to set a level of detail in a progressive mesh greater than the maximum available."
                Case D3DRMERR_REQUESTTOOSMALL
                    ErrorID = "D3DRMERR_REQUESTTOOSMALL"
                    Description = "Attempt was made to set the minimum rendering detail of a progressive mesh smaller than the detail in the base mesh (the minimum for rendering)."
                Case D3DRMERR_TEXTUREFORMATNOTFOUND
                    ErrorID = "D3DRMERR_TEXTUREFORMATNOTFOUND"
                    Description = "Texture format could not be found that meets the specified criteria and that the underlying Immediate Mode device supports."
                Case D3DRMERR_UNABLETOEXECUTE
                    ErrorID = "D3DRMERR_UNABLETOEXECUTE"
                    Description = "Unable to carry out procedure."
                Case DD_OK
                    ErrorID = "DD_OK"
                    Description = "Request completed successfully. Equivalent to D3DRM_OK."
                Case DDERR_INVALIDOBJECT
                    ErrorID = "DDERR_INVALIDOBJECT"
                    Description = "Received pointer that was an invalid object."
                Case DDERR_INVALIDPARAMS
                    ErrorID = "DDERR_INVALIDPARAMS"
                    Description = "One or more of the parameters passed to the method are incorrect."
                Case DDERR_NOTFOUND
                    ErrorID = "DDERR_NOTFOUND"
                    Description = "Requested item was not found."
                Case DDERR_NOTINITIALIZED
                    ErrorID = "DDERR_NOTINITIALIZED"
                    Description = "An attempt was made to call an interface method of an object before the object was initialized."
                Case DDERR_OUTOFMEMORY
                    ErrorID = "DDERR_OUTOFMEMORY"
                    Description = "DirectDraw does not have enough memory to perform the operation."
            End Select
        Case DirectDraw
            Select Case TempErr
                Case DD_OK
                    ErrorID = "DD_OK"
                    Description = "The request completed successfully."
                Case DDERR_ALREADYINITIALIZED
                    ErrorID = "DDERR_ALREADYINITIALIZED"
                    Description = "The object has already been initialized."
                Case DDERR_BLTFASTCANTCLIP
                    ErrorID = "DDERR_BLTFASTCANTCLIP"
                    Description = "A DirectDrawClipper object is attached to a source surface that has passed into a call to the DirectDrawSurface7.BltFast method."
                Case DDERR_CANNOTATTACHSURFACE
                    ErrorID = "DDERR_CANNOTATTACHSURFACE"
                    Description = "A surface cannot be attached to another requested surface."
                Case DDERR_CANNOTDETACHSURFACE
                    ErrorID = "DDERR_CANNOTDETACHSURFACE"
                    Description = "A surface cannot be detached from another requested surface."
                Case DDERR_CANTCREATEDC
                    ErrorID = "DDERR_CANTCREATEDC"
                    Description = "Windows cannot create any more device contexts (DCs), or a DC was requested for a palette-indexed surface when the surface had no palette and the display mode was not palette-indexed (in this case DirectDraw cannot select a proper palette into the DC)."
                Case DDERR_CANTDUPLICATE
                    ErrorID = "DDERR_CANTDUPLICATE"
                    Description = "Primary and 3-D surfaces, or surfaces that are implicitly created, cannot be duplicated."
                Case DDERR_CANTLOCKSURFACE
                    ErrorID = "DDERR_CANTLOCKSURFACE"
                    Description = "Access to this surface is refused because an attempt was made to lock the primary surface without DCI support."
                Case DDERR_CANTPAGELOCK
                    ErrorID = "DDERR_CANTPAGELOCK"
                    Description = "An attempt to page-lock a surface failed. Page lock does not work on a display-memory surface or an emulated primary surface."
                Case DDERR_CANTPAGEUNLOCK
                    ErrorID = "DDERR_CANTPAGEUNLOCK"
                    Description = "An attempt to page-unlock a surface failed. Page unlock does not work on a display-memory surface or an emulated primary surface."
                Case DDERR_CLIPPERISUSINGHWND
                    ErrorID = "DDERR_CLIPPERISUSINGHWND"
                    Description = "An attempt was made to set a clip list for a DirectDrawClipper object that is already monitoring a window handle."
                Case DDERR_COLORKEYNOTSET
                    ErrorID = "DDERR_COLORKEYNOTSET"
                    Description = "No source color key is specified for this operation."
                Case DDERR_CURRENTLYNOTAVAIL
                    ErrorID = "DDERR_CURRENTLYNOTAVAIL"
                    Description = "No support is currently available."
                Case DDERR_DCALREADYCREATED
                    ErrorID = "DDERR_DCALREADYCREATED"
                    Description = "A device context (DC) has already been returned for this surface. Only one DC can be retrieved for each surface."
                Case DDERR_DEVICEDOESNTOWNSURFACE
                    ErrorID = "DDERR_DEVICEDOESNTOWNSURFACE"
                    Description = "Surfaces created by one DirectDraw device cannot be used directly by another DirectDraw device."
                Case DDERR_DIRECTDRAWALREADYCREATED
                    ErrorID = "DDERR_DIRECTDRAWALREADYCREATED"
                    Description = "A DirectDraw object representing this driver has already been created for this process."
                Case DDERR_EXCEPTION
                    ErrorID = "DDERR_EXCEPTION"
                    Description = "An exception was encountered while performing the requested operation."
                Case DDERR_EXCLUSIVEMODEALREADYSET
                    ErrorID = "DDERR_EXCLUSIVEMODEALREADYSET"
                    Description = "An attempt was made to set the cooperative level when it was already set to exclusive."
                Case DDERR_EXPIRED
                    ErrorID = "DDERR_EXPIRED"
                    Description = "The data has expired and is therefore no longer valid."
                Case DDERR_GENERIC
                    ErrorID = "DDERR_GENERIC"
                    Description = "There is an undefined error condition."
                Case DDERR_HEIGHTALIGN
                    ErrorID = "DDERR_HEIGHTALIGN"
                    Description = "The height of the provided rectangle is not a multiple of the required alignment."
                Case DDERR_HWNDALREADYSET
                    ErrorID = "DDERR_HWNDALREADYSET"
                    Description = "The DirectDraw cooperative level window handle has already been set. It cannot be reset while the process has surfaces or palettes created."
                Case DDERR_HWNDSUBCLASSED
                    ErrorID = "DDERR_HWNDSUBCLASSED"
                    Description = "DirectDraw is prevented from restoring state because the DirectDraw cooperative level window handle has been subclassed."
                Case DDERR_IMPLICITLYCREATED
                    ErrorID = "DDERR_IMPLICITLYCREATED"
                    Description = "The surface cannot be restored because it is an implicitly created surface."
                Case DDERR_INCOMPATIBLEPRIMARY
                    ErrorID = "DDERR_INCOMPATIBLEPRIMARY"
                    Description = "The primary surface creation request does not match with the existing primary surface."
                Case DDERR_INVALIDCAPS
                    ErrorID = "DDERR_INVALIDCAPS"
                    Description = "One or more of the capability bits passed to the callback function are incorrect."
                Case DDERR_INVALIDCLIPLIST
                    ErrorID = "DDERR_INVALIDCLIPLIST"
                    Description = "DirectDraw does not support the provided clip list."
                Case DDERR_INVALIDDIRECTDRAWGUID
                    ErrorID = "DDERR_INVALIDDIRECTDRAWGUID"
                    Description = "The globally unique identifier (GUID) passed to the DirectX7.DirectDrawCreate function is not a valid DirectDraw driver identifier."
                Case DDERR_INVALIDMODE
                    ErrorID = "DDERR_INVALIDMODE"
                    Description = "DirectDraw does not support the requested mode."
                Case DDERR_INVALIDOBJECT
                    ErrorID = "DDERR_INVALIDOBJECT"
                    Description = "DirectDraw received a pointer that was an invalid DirectDraw object."
                Case DDERR_INVALIDPARAMS
                    ErrorID = "DDERR_INVALIDPARAMS"
                    Description = "One or more of the parameters passed to the method are incorrect."
                Case DDERR_INVALIDPIXELFORMAT
                    ErrorID = "DDERR_INVALIDPIXELFORMAT"
                    Description = "The pixel format was invalid as specified."
                Case DDERR_INVALIDPOSITION
                    ErrorID = "DDERR_INVALIDPOSITION"
                    Description = "The position of the overlay on the destination is no longer legal."
                Case DDERR_INVALIDRECT
                    ErrorID = "DDERR_INVALIDRECT"
                    Description = "The provided rectangle was invalid."
                Case DDERR_INVALIDSTREAM
                    ErrorID = "DDERR_INVALIDSTREAM"
                    Description = "The specified stream contains invalid data."
                Case DDERR_INVALIDSURFACETYPE
                    ErrorID = "DDERR_INVALIDSURFACETYPE"
                    Description = "The requested operation could not be performed because the surface was of the wrong type."
                Case DDERR_LOCKEDSURFACES
                    ErrorID = "DDERR_LOCKEDSURFACES"
                    Description = "One or more surfaces are locked."
                Case DDERR_MOREDATA
                    ErrorID = "DDERR_MOREDATA"
                    Description = "There is more data available than the specified buffer size can hold."
                Case DDERR_NO3D
                    ErrorID = "DDERR_NO3D"
                    Description = "No 3-D hardware or emulation is present."
                Case DDERR_NOALPHAHW
                    ErrorID = "DDERR_NOALPHAHW"
                    Description = "No alpha acceleration hardware is present or available."
                Case DDERR_NOBLTHW
                    ErrorID = "DDERR_NOBLTHW"
                    Description = "No blitter hardware is present."
                Case DDERR_NOCLIPLIST
                    ErrorID = "DDERR_NOCLIPLIST"
                    Description = "No clip list is available."
                Case DDERR_NOCLIPPERATTACHED
                    ErrorID = "DDERR_NOCLIPPERATTACHED"
                    Description = "No DirectDrawClipper object is attached to the surface object."
                Case DDERR_NOCOLORCONVHW
                    ErrorID = "DDERR_NOCOLORCONVHW"
                    Description = "No color-conversion hardware is present or available."
                Case DDERR_NOCOLORKEY
                    ErrorID = "DDERR_NOCOLORKEY"
                    Description = "The surface does not currently have a color key."
                Case DDERR_NOCOLORKEYHW
                    ErrorID = "DDERR_NOCOLORKEYHW"
                    Description = "There is no hardware support for the destination color key."
                Case DDERR_NOCOOPERATIVELEVELSET
                    ErrorID = "DDERR_NOCOOPERATIVELEVELSET"
                    Description = "A create function was called when the DirectDraw7.SetCooperativeLevel method had not been called."
                Case DDERR_NODC
                    ErrorID = "DDERR_NODC"
                    Description = "No DC has ever been created for this surface."
                Case DDERR_NODDROPSHW
                    ErrorID = "DDERR_NODDROPSHW"
                    Description = "No DirectDraw raster operation (ROP) hardware is available."
                Case DDERR_NODIRECTDRAWHW
                    ErrorID = "DDERR_NODIRECTDRAWHW"
                    Description = "Hardware-only DirectDraw object creation is not possible; the driver does not support any hardware."
                Case DDERR_NODIRECTDRAWSUPPORT
                    ErrorID = "DDERR_NODIRECTDRAWSUPPORT"
                    Description = "DirectDraw support is not possible with the current display driver."
                Case DDERR_NOEMULATION
                    ErrorID = "DDERR_NOEMULATION"
                    Description = "Software emulation is not available."
                Case DDERR_NOEXCLUSIVEMODE
                    ErrorID = "DDERR_NOEXCLUSIVEMODE"
                    Description = "The operation requires the application to have exclusive mode, but the application does not have exclusive mode."
                Case DDERR_NOFLIPHW
                    ErrorID = "DDERR_NOFLIPHW"
                    Description = "Flipping visible surfaces is not supported."
                Case DDERR_NOFOCUSWINDOW
                    ErrorID = "DDERR_NOFOCUSWINDOW"
                    Description = "An attempt was made to create or set a device window without first setting the focus window."
                Case DDERR_NOGDI
                    ErrorID = "DDERR_NOGDI"
                    Description = "No GDI is present."
                Case DDERR_NOHWND
                    ErrorID = "DDERR_NOHWND"
                    Description = "Clipper notification requires a window handle, or no window handle was previously set as the cooperative level window handle."
                Case DDERR_NOMIPMAPHW
                    ErrorID = "DDERR_NOMIPMAPHW"
                    Description = "No mipmap-capable texture mapping hardware is present or available."
                Case DDERR_NOMIRRORHW
                    ErrorID = "DDERR_NOMIRRORHW"
                    Description = "No mirroring hardware is present or available."
                Case DDERR_NONONLOCALVIDMEM
                    ErrorID = "DDERR_NONONLOCALVIDMEM"
                    Description = "An attempt was made to allocate nonlocal video memory from a device that does not support nonlocal video memory."
                Case DDERR_NOOPTIMIZEHW
                    ErrorID = "DDERR_NOOPTIMIZEHW"
                    Description = "The device does not support optimized surfaces."
                Case DDERR_NOOVERLAYHW
                    ErrorID = "DDERR_NOOVERLAYHW"
                    Description = "No overlay hardware is present or available."
                Case DDERR_NOPALETTEATTACHED
                    ErrorID = "DDERR_NOPALETTEATTACHED"
                    Description = "No palette object is attached to this surface."
                Case DDERR_NOPALETTEHW
                    ErrorID = "DDERR_NOPALETTEHW"
                    Description = "There is no hardware support for 16- or 256-color palettes."
                Case DDERR_NORASTEROPHW
                    ErrorID = "DDERR_NORASTEROPHW"
                    Description = "No appropriate raster operation hardware is present or available."
                Case DDERR_NOROTATIONHW
                    ErrorID = "DDERR_NOROTATIONHW"
                    Description = "No rotation hardware is present or available."
                Case DDERR_NOSTEREOHARDWARE
                    ErrorID = "DDERR_NOSTEREOHARDWARE"
                    Description = "No stereo hardware is present or available."
                Case DDERR_NOSTRETCHHW
                    ErrorID = "DDERR_NOSTRETCHHW"
                    Description = "There is no hardware support for stretching."
                Case DDERR_NOSURFACELEFT
                    ErrorID = "DDERR_NOSURFACELEFT"
                    Description = "No hardware is present that supports stereo surfaces."
                Case DDERR_NOT4BITCOLOR
                    ErrorID = "DDERR_NOT4BITCOLOR"
                    Description = "The DirectDrawSurface object is not using a 4-bit color palette, and the requested operation requires a 4-bit color palette."
                Case DDERR_NOT4BITCOLORINDEX
                    ErrorID = "DDERR_NOT4BITCOLORINDEX"
                    Description = "The DirectDrawSurface object is not using a 4-bit color index palette, and the requested operation requires a 4-bit color index palette."
                Case DDERR_NOT8BITCOLOR
                    ErrorID = "DDERR_NOT8BITCOLOR"
                    Description = "The DirectDrawSurface object is not using an 8-bit color palette, and the requested operation requires an 8-bit color palette."
                Case DDERR_NOTAOVERLAYSURFACE
                    ErrorID = "DDERR_NOTAOVERLAYSURFACE"
                    Description = "An overlay component was called for a non-overlay surface."
                Case DDERR_NOTEXTUREHW
                    ErrorID = "DDERR_NOTEXTUREHW"
                    Description = "No texture-mapping hardware is present or available."
                Case DDERR_NOTFLIPPABLE
                    ErrorID = "DDERR_NOTFLIPPABLE"
                    Description = "An attempt was made to flip a surface that cannot be flipped."
                Case DDERR_NOTFOUND
                    ErrorID = "DDERR_NOTFOUND"
                    Description = "The requested item was not found."
                Case DDERR_NOTINITIALIZED
                    ErrorID = "DDERR_NOTINITIALIZED"
                    Description = "An attempt was made to call an interface method of a DirectDraw object created by CoCreateInstance before the object was initialized."
                Case DDERR_NOTLOADED
                    ErrorID = "DDERR_NOTLOADED"
                    Description = "The surface is an optimized surface, but it has not yet been allocated any memory."
                Case DDERR_NOTLOCKED
                    ErrorID = "DDERR_NOTLOCKED"
                    Description = "An attempt was made to unlock a surface that was not locked."
                Case DDERR_NOTPAGELOCKED
                    ErrorID = "DDERR_NOTPAGELOCKED"
                    Description = "An attempt was made to page-unlock a surface with no outstanding page locks."
                Case DDERR_NOTPALETTIZED
                    ErrorID = "DDERR_NOTPALETTIZED"
                    Description = "The surface being used is not a palette-based surface."
                Case DDERR_NOVSYNCHW
                    ErrorID = "DDERR_NOVSYNCHW"
                    Description = "There is no hardware support for vertical blank synchronized operations."
                Case DDERR_NOZBUFFERHW
                    ErrorID = "DDERR_NOZBUFFERHW"
                    Description = "There is no hardware support for z-buffers."
                Case DDERR_NOZOVERLAYHW
                    ErrorID = "DDERR_NOZOVERLAYHW"
                    Description = "The hardware does not support z-ordering of overlays."
                Case DDERR_OUTOFCAPS
                    ErrorID = "DDERR_OUTOFCAPS"
                    Description = "The hardware needed for the requested operation has already been allocated."
                Case DDERR_OUTOFMEMORY
                    ErrorID = "DDERR_OUTOFMEMORY"
                    Description = "DirectDraw does not have enough memory to perform the operation."
                Case DDERR_OUTOFVIDEOMEMORY
                    ErrorID = "DDERR_OUTOFVIDEOMEMORY"
                    Description = "DirectDraw does not have enough display memory to perform the operation."
                Case DDERR_OVERLAPPINGRECTS
                    ErrorID = "DDERR_OVERLAPPINGRECTS"
                    Description = "The source and destination rectangles are on the same surface and overlap each other."
                Case DDERR_OVERLAYCANTCLIP
                    ErrorID = "DDERR_OVERLAYCANTCLIP"
                    Description = "The hardware does not support clipped overlays."
                Case DDERR_OVERLAYCOLORKEYONLYONEACTIVE
                    ErrorID = "DDERR_OVERLAYCOLORKEYONLYONEACTIVE"
                    Description = "An attempt was made to have more than one color key active on an overlay."
                Case DDERR_OVERLAYNOTVISIBLE
                    ErrorID = "DDERR_OVERLAYNOTVISIBLE"
                    Description = "The method was called on a hidden overlay."
                Case DDERR_PALETTEBUSY
                    ErrorID = "DDERR_PALETTEBUSY"
                    Description = "Access to this palette is refused because the palette is locked by another thread."
                Case DDERR_PRIMARYSURFACEALREADYEXISTS
                    ErrorID = "DDERR_PRIMARYSURFACEALREADYEXISTS"
                    Description = "This process has already created a primary surface."
                Case DDERR_REGIONTOOSMALL
                    ErrorID = "DDERR_REGIONTOOSMALL"
                    Description = "The region passed to the DirectDrawClipper.GetClipList method is too small."
                Case DDERR_SURFACEALREADYATTACHED
                    ErrorID = "DDERR_SURFACEALREADYATTACHED"
                    Description = "An attempt was made to attach a surface to another surface to which it is already attached."
                Case DDERR_SURFACEALREADYDEPENDENT
                    ErrorID = "DDERR_SURFACEALREADYDEPENDENT"
                    Description = "An attempt was made to make a surface a dependency of another surface on which it is already dependent."
                Case DDERR_SURFACEBUSY
                    ErrorID = "DDERR_SURFACEBUSY"
                    Description = "Access to the surface is refused because the surface is locked by another thread."
                Case DDERR_SURFACEISOBSCURED
                    ErrorID = "DDERR_SURFACEISOBSCURED"
                    Description = "Access to the surface is refused because the surface is obscured."
                Case DDERR_SURFACELOST
                    ErrorID = "DDERR_SURFACELOST"
                    Description = "Access to the surface is refused because the surface memory is gone. Call the DirectDrawSurface7.Restore method on this surface to restore the memory associated with it."
                Case DDERR_SURFACENOTATTACHED
                    ErrorID = "DDERR_SURFACENOTATTACHED"
                    Description = "The requested surface is not attached."
                Case DDERR_TOOBIGHEIGHT
                    ErrorID = "DDERR_TOOBIGHEIGHT"
                    Description = "The height requested by DirectDraw is too large."
                Case DDERR_TOOBIGSIZE
                    ErrorID = "DDERR_TOOBIGSIZE"
                    Description = "The size requested by DirectDraw is too large. However, the individual height and width are valid sizes."
                Case DDERR_TOOBIGWIDTH
                    ErrorID = "DDERR_TOOBIGWIDTH"
                    Description = "The width requested by DirectDraw is too large."
                Case DDERR_UNSUPPORTED
                    ErrorID = "DDERR_UNSUPPORTED"
                    Description = "The operation is not supported."
                Case DDERR_UNSUPPORTEDFORMAT
                    ErrorID = "DDERR_UNSUPPORTEDFORMAT"
                    Description = "The FourCC format requested is not supported by DirectDraw."
                Case DDERR_UNSUPPORTEDMASK
                    ErrorID = "DDERR_UNSUPPORTEDMASK"
                    Description = "The bitmask in the pixel format requested is not supported by DirectDraw."
                Case DDERR_UNSUPPORTEDMODE
                    ErrorID = "DDERR_UNSUPPORTEDMODE"
                    Description = "The display is currently in an unsupported mode."
                Case DDERR_VERTICALBLANKINPROGRESS
                    ErrorID = "DDERR_VERTICALBLANKINPROGRESS"
                    Description = "A vertical blank is in progress."
                Case DDERR_VIDEONOTACTIVE
                    ErrorID = "DDERR_VIDEONOTACTIVE"
                    Description = "The video port is not active."
                Case DDERR_WASSTILLDRAWING
                    ErrorID = "DDERR_WASSTILLDRAWING"
                    Description = "The previous blit operation that is transferring information to or from this surface is incomplete."
                Case DDERR_WRONGMODE
                    ErrorID = "DDERR_WRONGMODE"
                    Description = "This surface cannot be restored because it was created in a different mode."
                Case DDERR_XALIGN
                    ErrorID = "DDERR_XALIGN"
                    Description = "The provided rectangle was not horizontally aligned on a required boundary."
                'Case E_INVALIDINTERFACE
                '    ErrorID = "E_INVALIDINTERFACE"
                '    Description = "The specified interface is invalid or does not exist."
                'Case E_OUTOFMEMORY
                '    ErrorID = "E_OUTOFMEMORY"
                '    Description = "Not enough free memory to complete the method."
            End Select
        Case DirectInput
            Select Case TempErr
                Case DI_BUFFEROVERFLOW
                    ErrorID = "DI_BUFFEROVERFLOW"
                    Description = "The input buffer overflowed and data was lost."
                Case DIERR_ACQUIRED
                    ErrorID = "DIERR_ACQUIRED"
                    Description = "The operation cannot be performed while the device is acquired."
                Case DIERR_ALREADYINITIALIZED
                    ErrorID = "DIERR_ALREADYINITIALIZED"
                    Description = "This object is already initialized"
                Case DIERR_BADDRIVERVER
                    ErrorID = "DIERR_BADDRIVERVER"
                    Description = "The object could not be created due to an incompatible driver version or mismatched or incomplete driver components."
                Case DIERR_BETADIRECTINPUTVERSION
                    ErrorID = "DIERR_BETADIRECTINPUTVERSION"
                    Description = "The application was written for an unsupported prerelease version of DirectInput."
                Case DIERR_DEVICEFULL
                    ErrorID = "DIERR_DEVICEFULL"
                    Description = "The device is full."
                Case DIERR_DEVICENOTREG
                    ErrorID = "DIERR_DEVICENOTREG"
                    Description = "The device or device instance is not registered with DirectInput. This value is equal to the REGDB_E_CLASSNOTREG standard COM return value."
                Case DIERR_EFFECTPLAYING
                    ErrorID = "DIERR_EFFECTPLAYING"
                    Description = "The parameters were updated in memory but were not downloaded to the device because the device does not support updating an effect while it is still playing."
                Case DIERR_HASEFFECTS
                    ErrorID = "DIERR_HASEFFECTS"
                    Description = "The device cannot be reinitialized because there are still effects attached to it."
                Case DIERR_GENERIC
                    ErrorID = "DIERR_GENERIC"
                    Description = "An undetermined error occurred inside the DirectInput subsystem. This value is equal to the E_FAIL standard COM return value."
                Case DIERR_HANDLEEXISTS
                    ErrorID = "DIERR_HANDLEEXISTS"
                    Description = "The device already has an event notification associated with it. This value is equal to the E_ACCESSDENIED standard COM return value."
                Case DIERR_INCOMPLETEEFFECT
                    ErrorID = "DIERR_INCOMPLETEEFFECT"
                    Description = "The effect could not be downloaded because essential information is missing. For example, no axes have been associated with the effect, or no type-specific information has been supplied."
                Case DIERR_INPUTLOST
                    ErrorID = "DIERR_INPUTLOST"
                    Description = "Access to the input device has been lost. It must be reacquired."
                Case DIERR_INVALIDHANDLE
                    ErrorID = "DIERR_INVALIDHANDLE"
                    Description = "An invalid window handle was passed to the method."
                Case DIERR_INVALIDPARAM
                    ErrorID = "DIERR_INVALIDPARAM"
                    Description = "An invalid parameter was passed to the returning function, or the object was not in a state that permitted the function to be called. This value is equal to the E_INVALIDARG standard COM return value."
                Case DIERR_MOREDATA
                    ErrorID = "DIERR_MOREDATA"
                    Description = "Not all the requested information fitted into the buffer."
                Case DIERR_NOAGGREGATION
                    ErrorID = "DIERR_NOAGGREGATION"
                    Description = "This object does not support aggregation."
                Case DIERR_NOINTERFACE
                    ErrorID = "DIERR_NOINTERFACE"
                    Description = "The specified interface is not supported by the object. This value is equal to the E_NOINTERFACE standard COM return value."
                Case DIERR_NOTACQUIRED
                    ErrorID = "DIERR_NOTACQUIRED"
                    Description = "The operation cannot be performed unless the device is acquired."
                Case DIERR_NOTBUFFERED
                    ErrorID = "DIERR_NOTBUFFERED"
                    Description = "The device is not buffered. Set the DIPROP_BUFFERSIZE property to enable buffering."
                Case DIERR_NOTDOWNLOADED
                    ErrorID = "DIERR_NOTDOWNLOADED"
                    Description = "The effect is not downloaded."
                Case DIERR_NOTEXCLUSIVEACQUIRED
                    ErrorID = "DIERR_NOTEXCLUSIVEACQUIRED"
                    Description = "The operation cannot be performed unless the device is acquired in DISCL_EXCLUSIVE mode."
                Case DIERR_NOTINITIALIZED
                    ErrorID = "DIERR_NOTINITIALIZED"
                    Description = "The object has not been initialized."
                Case DIERR_NOTFOUND
                    ErrorID = "DIERR_NOTFOUND"
                    Description = "The requested object does not exist."
                Case DIERR_OBJECTNOTFOUND
                    ErrorID = "DIERR_OBJECTNOTFOUND"
                    Description = "The requested object does not exist."
                Case DIERR_OLDDIRECTINPUTVERSION
                    ErrorID = "DIERR_OLDDIRECTINPUTVERSION"
                    Description = "The application requires a newer version of DirectInput."
                Case DIERR_OTHERAPPHASPRIO
                    ErrorID = "DIERR_OTHERAPPHASPRIO"
                    Description = "Another application has a higher priority level, preventing this call from succeeding. This value is equal to the E_ACCESSDENIED standard COM return value. This error can be returned when an application has only foreground access to a device but is attempting to acquire the device while in the background."
                Case DIERR_OUTOFMEMORY
                    ErrorID = "DIERR_OUTOFMEMORY"
                    Description = "The DirectInput subsystem couldn't allocate sufficient memory to complete the call. This value is equal to the E_OUTOFMEMORY standard COM return value."
                Case DIERR_READONLY
                    ErrorID = "DIERR_READONLY"
                    Description = "The specified property cannot be changed. This value is equal to the E_ACCESSDENIED standard COM return value."
                'Case DIERR_REPORTFULL
                '    ErrorID = "DIERR_REPORTFULL"
                '    Description = "More information was requested to be sent than can be sent to the device."
                'Case DIERR_UNPLUGGED
                '    ErrorID = "DIERR_UNPLUGGED"
                '    Description = "The operation could not be completed because the device is not plugged in."
                Case DIERR_UNSUPPORTED
                    ErrorID = "DIERR_UNSUPPORTED"
                    Description = "The function called is not supported at this time. This value is equal to the E_NOTIMPL standard COM return value."
                Case E_PENDING
                    ErrorID = "E_PENDING"
                    Description = "Data is not yet available."
            End Select
        Case DirectMusic
            Select Case TempErr
                Case DMUS_E_ALL_TRACKS_FAILED
                    ErrorID = "DMUS_E_ALL_TRACKS_FAILED"
                    Description = "The segment object was unable to load all tracks from the IStream object data, perhaps because of errors in the stream or because the tracks are incorrectly registered on the client."
                Case DMUS_E_ALREADY_ACTIVATED
                    ErrorID = "DMUS_E_ALREADY_ACTIVATED"
                    Description = "The port has been activated, and the parameter cannot be changed."
                Case DMUS_E_ALREADY_DOWNLOADED
                    ErrorID = "DMUS_E_ALREADY_DOWNLOADED"
                    Description = "Buffer has already been downloaded."
                Case DMUS_E_ALREADY_EXISTS
                    ErrorID = "DMUS_E_ALREADY_EXISTS"
                    Description = "The tool is already contained in the graph. You must create a new instance."
                Case DMUS_E_ALREADY_INITED
                    ErrorID = "DMUS_E_ALREADY_INITED"
                    Description = "The object has already been initialized."
                Case DMUS_E_ALREADY_LOADED
                    ErrorID = "DMUS_E_ALREADY_LOADED"
                    Description = "The DLS collection is already open."
                Case DMUS_E_ALREADY_SENT
                    ErrorID = "DMUS_E_ALREADY_SENT"
                    Description = "The message has already been sent."
                Case DMUS_E_ALREADYCLOSED
                    ErrorID = "DMUS_E_ALREADYCLOSED"
                    Description = "The port is not open."
                Case DMUS_E_ALREADYOPEN
                    ErrorID = "DMUS_E_ALREADYOPEN"
                    Description = "The port was already opened."
                Case DMUS_E_BADARTICULATION
                    ErrorID = "DMUS_E_BADARTICULATION"
                    Description = "Invalid articulation chunk in DLS collection."
                Case DMUS_E_BADINSTRUMENT
                    ErrorID = "DMUS_E_BADINSTRUMENT"
                    Description = "Invalid instrument chunk in DLS collection."
                Case DMUS_E_BADOFFSETTABLE
                    ErrorID = "DMUS_E_BADOFFSETTABLE"
                    Description = "Offset table has errors."
                Case DMUS_E_BADWAVE
                    ErrorID = "DMUS_E_BADWAVE"
                    Description = "Corrupt wave header."
                Case DMUS_E_BADWAVELINK
                    ErrorID = "DMUS_E_BADWAVELINK"
                    Description = "Wave-link chunk in DLS collection points to an invalid wave."
                Case DMUS_E_BUFFER_EMPTY
                    ErrorID = "DMUS_E_BUFFER_EMPTY"
                    Description = "There is no data in the buffer."
                Case DMUS_E_BUFFER_FULL
                    ErrorID = "DMUS_E_BUFFER_FULL"
                    Description = "The specified number of bytes exceeds the maximum buffer size."
                Case DMUS_E_BUFFERNOTAVAILABLE
                    ErrorID = "DMUS_E_BUFFERNOTAVAILABLE"
                    Description = "The buffer is not available for download."
                Case DMUS_E_BUFFERNOTSET
                    ErrorID = "DMUS_E_BUFFERNOTSET"
                    Description = "No buffer was prepared for the data."
                Case DMUS_E_CANNOT_OPEN_PORT
                    ErrorID = "DMUS_E_CANNOT_OPEN_PORT"
                    Description = "The default system port could not be opened."
                Case DMUS_E_DEVICE_IN_USE
                    ErrorID = "DMUS_E_DEVICE_IN_USE"
                    Description = "Device is already in use (possibly by a non-DirectMusic client) and cannot be opened again."
                Case DMUS_E_DMUSIC_RELEASED
                    ErrorID = "DMUS_E_DMUSIC_RELEASED"
                    Description = "Operation cannot be performed because the final instance of the DirectMusic object was released. Ports cannot be used after final release of the DirectMusic object."
                Case DMUS_E_DRIVER_FAILED
                    ErrorID = "DMUS_E_DRIVER_FAILED"
                    Description = "An unexpected error was returned from a device driver, indicating possible failure of the driver or hardware."
                Case DMUS_E_DSOUND_ALREADY_SET
                    ErrorID = "DMUS_E_DSOUND_ALREADY_SET"
                    Description = "A DirectSound object has already been set."
                Case DMUS_E_DSOUND_NOT_SET
                    ErrorID = "DMUS_E_DSOUND_NOT_SET"
                    Description = "Port could not be created because no DirectSound object has been specified."
                Case DMUS_E_FAIL
                    ErrorID = "DMUS_E_FAIL"
                    Description = "The method did not succeed."
                Case DMUS_E_GET_UNSUPPORTED
                    ErrorID = "DMUS_E_GET_UNSUPPORTED"
                    Description = "Getting the parameter is not supported."
                Case DMUS_E_INSUFFICIENTBUFFER
                    ErrorID = "DMUS_E_INSUFFICIENTBUFFER"
                    Description = "Buffer is not large enough for the requested operation."
                Case DMUS_E_INVALIDARG
                    ErrorID = "DMUS_E_INVALIDARG"
                    Description = "Invalid argument."
                Case DMUS_E_INVALID_BAND
                    ErrorID = "DMUS_E_INVALID_BAND"
                    Description = "File does not contain a valid band."
                Case DMUS_E_INVALID_DOWNLOADID
                    ErrorID = "DMUS_E_INVALID_DOWNLOADID"
                    Description = "Invalid download identifier was used in the process of creating a download buffer."
                Case DMUS_E_INVALID_EVENT
                    ErrorID = "DMUS_E_INVALID_EVENT"
                    Description = "The event either is not a valid MIDI message or makes use of running status, and cannot be packed into the buffer."
                Case DMUS_E_INVALIDBUFFER
                    ErrorID = "DMUS_E_INVALIDBUFFER"
                    Description = "Invalid DirectSound buffer was handed to port."
                Case DMUS_E_INVALIDFILE
                    ErrorID = "DMUS_E_INVALIDFILE"
                    Description = "Not a valid file."
                Case DMUS_E_INVALIDPATCH
                    ErrorID = "DMUS_E_INVALIDPATCH"
                    Description = "No instrument in the collection matches the patch number."
                Case DMUS_E_INVALIDPOS
                    ErrorID = "DMUS_E_INVALIDPOS"
                    Description = "Error reading wave data from a DLS collection. Indicates a bad file."
                Case DMUS_E_LOADER_BADPATH
                    ErrorID = "DMUS_E_LOADER_BADPATH"
                    Description = "The file path is invalid."
                Case DMUS_E_LOADER_FAILEDCREATE
                    ErrorID = "DMUS_E_LOADER_FAILEDCREATE"
                    Description = "Object could not be found or created."
                Case DMUS_E_LOADER_FAILEDOPEN
                    ErrorID = "DMUS_E_LOADER_FAILEDOPEN"
                    Description = "File open failed because the file does not exist or is locked."
                Case DMUS_E_LOADER_FORMATNOTSUPPORTED
                    ErrorID = "DMUS_E_LOADER_FORMATNOTSUPPORTED"
                    Description = "The object cannot be loaded because the data format is not supported."
                Case DMUS_E_LOADER_OBJECTNOTFOUND
                    ErrorID = "DMUS_E_LOADER_OBJECTNOTFOUND"
                    Description = "The object was not found."
                Case DMUS_E_NO_MASTER_CLOCK
                    ErrorID = "DMUS_E_NO_MASTER_CLOCK"
                    Description = "There is no master clock in the performance. Be sure to call the DirectMusicPerformance.Init method."
                Case DMUS_E_NOINTERFACE
                    ErrorID = "DMUS_E_NOINTERFACE"
                    Description = "No object interface is available."
                Case DMUS_E_NOT_DOWNLOADED_TO_PORT
                    ErrorID = "DMUS_E_NOT_DOWNLOADED_TO_PORT"
                    Description = "The object cannot be unloaded because it is not present on the port."
                Case DMUS_E_NOT_FOUND
                    ErrorID = "DMUS_E_NOT_FOUND"
                    Description = "The requested item is not contained by the object."
                Case DMUS_E_NOT_INIT
                    ErrorID = "DMUS_E_NOT_INIT"
                    Description = "A required object is not initialized or failed to initialize."
                Case DMUS_E_NOTADLSCOL
                    ErrorID = "DMUS_E_NOTADLSCOL"
                    Description = "The object being loaded is not a valid DLS collection."
                Case DMUS_E_NOTIMPL
                    ErrorID = "DMUS_E_NOTIMPL"
                    Description = "The method is not implemented. This value can be returned if a driver does not support a feature necessary for the operation."
                Case DMUS_E_OUT_OF_RANGE
                    ErrorID = "DMUS_E_OUT_OF_RANGE"
                    Description = "The requested time is outside the range of the segment."
                Case DMUS_E_OUTOFMEMORY
                    ErrorID = "DMUS_E_OUTOFMEMORY"
                    Description = "Insufficient memory to complete task."
                Case DMUS_E_PORT_NOT_RENDER
                    ErrorID = "DMUS_E_PORT_NOT_RENDER"
                    Description = "Not an output port."
                Case DMUS_E_PORTS_OPEN
                    ErrorID = "DMUS_E_PORTS_OPEN"
                    Description = "The requested operation cannot be performed while there are instantiated ports in any process in the system."
                Case DMUS_E_SEGMENT_INIT_FAILED
                    ErrorID = "DMUS_E_SEGMENT_INIT_FAILED"
                    Description = "Segment initialization failed, probably because of a critical memory situation."
                Case DMUS_E_SET_UNSUPPORTED
                    ErrorID = "DMUS_E_SET_UNSUPPORTED"
                    Description = "Setting the parameter is not supported."
                Case DMUS_E_TIME_PAST
                    ErrorID = "DMUS_E_TIME_PAST"
                    Description = "The time requested is in the past."
                Case DMUS_E_TRACK_NOT_FOUND
                    ErrorID = "DMUS_E_TRACK_NOT_FOUND"
                    Description = "There is no track of the requested type."
                Case DMUS_E_TYPE_DISABLED
                    ErrorID = "DMUS_E_TYPE_DISABLED"
                    Description = "Parameter is unavailable because it has been disabled."
                Case DMUS_E_TYPE_UNSUPPORTED
                    ErrorID = "DMUS_E_TYPE_UNSUPPORTED"
                    Description = "Parameter is unsupported on this track."
                Case DMUS_E_UNKNOWN_PROPERTY
                    ErrorID = "DMUS_E_UNKNOWN_PROPERTY"
                    Description = "The property set or item is not implemented by this port."
                Case DMUS_E_UNSUPPORTED_STREAM
                    ErrorID = "DMUS_E_UNSUPPORTED_STREAM"
                    Description = "The stream does not contain data supported by the loading object."
            End Select
        Case DirectPlay
            Select Case TempErr
                Case DP_OK
                    ErrorID = "DP_OK"
                    Description = "The request completed successfully."
                Case DPERR_ABORTED
                    ErrorID = "DPERR_ABORTED"
                    Description = "The operation was canceled before it could be completed."
                Case DPERR_ACCESSDENIED
                    ErrorID = "DPERR_ACCESSDENIED"
                    Description = "The session is full, or an incorrect password was supplied."
                Case DPERR_ACTIVEPLAYERS
                    ErrorID = "DPERR_ACTIVEPLAYERS"
                    Description = "The requested operation cannot be performed because there are existing active players."
                Case DPERR_ALREADYINITIALIZED
                    ErrorID = "DPERR_ALREADYINITIALIZED"
                    Description = "This object is already initialized."
                Case DPERR_APPNOTSTARTED
                    ErrorID = "DPERR_APPNOTSTARTED"
                    Description = "The application has not been started yet."
                Case DPERR_AUTHENTICATIONFAILED
                    ErrorID = "DPERR_AUTHENTICATIONFAILED"
                    Description = "The password or credentials supplied could not be authenticated."
                Case DPERR_BUFFERTOOLARGE
                    ErrorID = "DPERR_BUFFERTOOLARGE"
                    Description = "The data buffer is too large to store."
                Case DPERR_BUFFERTOOSMALL
                    ErrorID = "DPERR_BUFFERTOOSMALL"
                    Description = "The supplied buffer is not large enough to contain the requested data."
                Case DPERR_BUSY
                    ErrorID = "DPERR_BUSY"
                    Description = "A message cannot be sent because the transmission medium is busy."
                Case DPERR_CANCELFAILED
                    ErrorID = "DPERR_CANCELFAILED"
                    Description = "The message could not be canceled, possibly because it is a group message that has already been to sent to one or more members of the group."
                Case DPERR_CANCELLED
                    ErrorID = "DPERR_CANCELLED"
                    Description = "The operation was canceled."
                Case DPERR_CANNOTCREATESERVER
                    ErrorID = "DPERR_CANNOTCREATESERVER"
                    Description = "The server cannot be created for the new session."
                Case DPERR_CANTADDPLAYER
                    ErrorID = "DPERR_CANTADDPLAYER"
                    Description = "The player cannot be added to the session."
                Case DPERR_CANTCREATEGROUP
                    ErrorID = "DPERR_CANTCREATEGROUP"
                    Description = "A new group cannot be created."
                Case DPERR_CANTCREATEPLAYER
                    ErrorID = "DPERR_CANTCREATEPLAYER"
                    Description = "A new player cannot be created."
                Case DPERR_CANTCREATEPROCESS
                    ErrorID = "DPERR_CANTCREATEPROCESS"
                    Description = "Cannot start the application."
                Case DPERR_CANTCREATESESSION
                    ErrorID = "DPERR_CANTCREATESESSION"
                    Description = "A new session cannot be created."
                Case DPERR_CANTLOADCAPI
                    ErrorID = "DPERR_CANTLOADCAPI"
                    Description = "No credentials were supplied and the CryptoAPI package (CAPI) to use for cryptography services cannot be loaded."
                Case DPERR_CANTLOADSECURITYPACKAGE
                    ErrorID = "DPERR_CANTLOADSECURITYPACKAGE"
                    Description = "The software security package cannot be loaded."
                Case DPERR_CANTLOADSSPI
                    ErrorID = "DPERR_CANTLOADSSPI"
                    Description = "No credentials were supplied, and the Security Support Provider Interface (SSPI) that will prompt for credentials cannot be loaded."
                Case DPERR_CAPSNOTAVAILABLEYET
                    ErrorID = "DPERR_CAPSNOTAVAILABLEYET"
                    Description = "The capabilities of the DirectPlay object have not been determined yet. This error will occur if the DirectPlay object is implemented on a connectivity solution that requires polling to determine available bandwidth and latency."
                Case DPERR_CONNECTING
                    ErrorID = "DPERR_CONNECTING"
                    Description = "The method is in the process of connecting to the network. The application should keep using the method until it returns DP_OK, indicating successful completion, or until it returns a different error."
                Case DPERR_CONNECTIONLOST
                    ErrorID = "DPERR_CONNECTIONLOST"
                    Description = "The service provider connection was reset while data was being sent."
                Case DPERR_ENCRYPTIONFAILED
                    ErrorID = "DPERR_ENCRYPTIONFAILED"
                    Description = "The requested information could not be digitally encrypted. Encryption is used for message privacy. This error is only relevant in a secure session."
                Case DPERR_EXCEPTION
                    ErrorID = "DPERR_EXCEPTION"
                    Description = "An exception occurred when processing the request."
                Case DPERR_GENERIC
                    ErrorID = "DPERR_GENERIC"
                    Description = "An undefined error condition occurred."
                Case DPERR_INVALIDFLAGS
                    ErrorID = "DPERR_INVALIDFLAGS"
                    Description = "The flags passed to this method are invalid."
                Case DPERR_INVALIDGROUP
                    ErrorID = "DPERR_INVALIDGROUP"
                    Description = "The group ID is not recognized as a valid group ID for this game session."
                Case DPERR_INVALIDINTERFACE
                    ErrorID = "DPERR_INVALIDINTERFACE"
                    Description = "The interface parameter is invalid."
                Case DPERR_INVALIDOBJECT
                    ErrorID = "DPERR_INVALIDOBJECT"
                    Description = "The DirectPlay object is invalid."
                Case DPERR_INVALIDPARAMS
                    ErrorID = "DPERR_INVALIDPARAMS"
                    Description = "One or more of the parameters passed to the method are invalid."
                Case DPERR_INVALIDPASSWORD
                    ErrorID = "DPERR_INVALIDPASSWORD"
                    Description = "An invalid password was supplied when attempting to join a session that requires a password."
                Case DPERR_INVALIDPLAYER
                    ErrorID = "DPERR_INVALIDPLAYER"
                    Description = "The player ID is not recognized as a valid player ID for this game session."
                Case DPERR_INVALIDPRIORITY
                    ErrorID = "DPERR_INVALIDPRIORITY"
                    Description = "The specified priority is not within the range of allowed priorities, which is inclusively 0-65535."
                Case DPERR_LOGONDENIED
                    ErrorID = "DPERR_LOGONDENIED"
                    Description = "The session could not be opened because credentials are required, and either no credentials were supplied, or the credentials were invalid."
                Case DPERR_NOCAPS
                    ErrorID = "DPERR_NOCAPS"
                    Description = "The communication link that DirectPlay is attempting to use is not capable of this function."
                Case DPERR_NOCONNECTION
                    ErrorID = "DPERR_NOCONNECTION"
                    Description = "No communication link was established."
                Case DPERR_NOINTERFACE
                    ErrorID = "DPERR_NOINTERFACE"
                    Description = "The interface is not supported."
                Case DPERR_NOMESSAGES
                    ErrorID = "DPERR_NOMESSAGES"
                    Description = "There are no messages in the receive queue."
                Case DPERR_NONAMESERVERFOUND
                    ErrorID = "DPERR_NONAMESERVERFOUND"
                    Description = "No name server (host) could be found or created. A host must exisglosate a player."
                Case DPERR_NONEWPLAYERS
                    ErrorID = "DPERR_NONEWPLAYERS"
                    Description = "The session is not accepting any new players."
                Case DPERR_NOPLAYERS
                    ErrorID = "DPERR_NOPLAYERS"
                    Description = "There are no active players in the session."
                Case DPERR_NOSESSIONS
                    ErrorID = "DPERR_NOSESSIONS"
                    Description = "There are no sessions for which this method can be called."
                Case DPERR_NOTLOBBIED
                    ErrorID = "DPERR_NOTLOBBIED"
                    Description = "Returned by the DirectPlayLobby3.Connect method if the application was not started by using the DirectPlayLobby3.RunApplication method, or if there is no DirectPlayLobbyConnection interface currently initialized for this DirectPlayLobby object."
                Case DPERR_NOTLOGGEDIN
                    ErrorID = "DPERR_NOTLOGGEDIN"
                    Description = "An action cannot be performed because a player or client application is not logged on. Returned by the DirectPlay4.Send method when the client application tries to send a secure message without being logged on."
                Case DPERR_OUTOFMEMORY
                    ErrorID = "DPERR_OUTOFMEMORY"
                    Description = "There is insufficient memory to perform the requested operation."
                Case DPERR_PENDING
                    ErrorID = "DPERR_PENDING"
                    Description = "Not an error, this return indicates that an asynchronous send has reached the point where it is successfully queued. See SendEx for more information."
                Case DPERR_PLAYERLOST
                    ErrorID = "DPERR_PLAYERLOST"
                    Description = "A player has lost the connection to the session."
                Case DPERR_SENDTOOBIG
                    ErrorID = "DPERR_SENDTOOBIG"
                    Description = "The message being sent by the DirectPlay4.Send method is too large."
                Case DPERR_SESSIONLOST
                    ErrorID = "DPERR_SESSIONLOST"
                    Description = "The connection to the session has been lost."
                Case DPERR_SIGNFAILED
                    ErrorID = "DPERR_SIGNFAILED"
                    Description = "The requested information could not be digitally signed. Digital signatures are used to establish the authenticity of messages."
                Case DPERR_TIMEOUT
                    ErrorID = "DPERR_TIMEOUT"
                    Description = "The operation could not be completed in the specified time."
                Case DPERR_UNAVAILABLE
                    ErrorID = "DPERR_UNAVAILABLE"
                    Description = "The requested function is not available at this time."
                Case DPERR_UNINITIALIZED
                    ErrorID = "DPERR_UNINITIALIZED"
                    Description = "The requested object has not been initialized."
                Case DPERR_UNKNOWNAPPLICATION
                    ErrorID = "DPERR_UNKNOWNAPPLICATION"
                    Description = "An unknown application was specified."
                Case DPERR_UNKNOWNMESSAGE
                    ErrorID = "DPERR_UNKNOWNMESSAGE"
                    Description = "The message ID isn't valid. Returned from DirectPlay4.CancelMessage if the ID of the message to be canceled is invalid."
                Case DPERR_UNSUPPORTED
                    ErrorID = "DPERR_UNSUPPORTED"
                    Description = "The function or feature is not available in this implementation or on this service provider. Returned from DirectPlay4.SetGroupConnectionSettings if this method is called from a session that is not a lobby session. Returned from DirectPlay4.SendEx if the priority or time-out is set, and these are not supported by the service provider and DirectPlay protocol is not on. Returned from DirectPlay4.GetMessageQueue if you check the send queue and it is not supported by the service provider and DirectPlay protocol is not on."
                Case DPERR_USERCANCEL
                    ErrorID = "DPERR_USERCANCEL"
                    Description = "Can be returned in two ways. 1) The user canceled the connection process during a call to the DirectPlay4.Open method. 2) The user clicked Cancel in one of the DirectPlay service provider dialog boxes during a call to DirectPlay4.GetDPEnumSessions."
            End Select
        Case DirectSound
            Select Case TempErr
                Case DS_OK
                    ErrorID = "DS_OK"
                    Description = "The request completed successfully."
                Case DSERR_ALLOCATED
                    ErrorID = "DSERR_ALLOCATED"
                    Description = "The request failed because resources, such as a priority level, were already in use by another caller."
                Case DSERR_ALREADYINITIALIZED
                    ErrorID = "DSERR_ALREADYINITIALIZED"
                    Description = "The object is already initialized."
                Case DSERR_BADFORMAT
                    ErrorID = "DSERR_BADFORMAT"
                    Description = "The specified wave format is not supported."
                Case DSERR_BUFFERLOST
                    ErrorID = "DSERR_BUFFERLOST"
                    Description = "The buffer memory has been lost and must be restored."
                Case DSERR_CONTROLUNAVAIL
                    ErrorID = "DSERR_CONTROLUNAVAIL"
                    Description = "The control (volume, pan, and so forth) requested by the caller is not available."
                Case DSERR_GENERIC
                    ErrorID = "DSERR_GENERIC"
                    Description = "An undetermined error occurred inside the DirectSound subsystem."
                Case DSERR_INVALIDCALL
                    ErrorID = "DSERR_INVALIDCALL"
                    Description = "This function is not valid for the current state of this object."
                Case DSERR_INVALIDPARAM
                    ErrorID = "DSERR_INVALIDPARAM"
                    Description = "An invalid parameter was passed to the returning function."
                'Case DSERR_NOAGGREGATION
                '    ErrorID = "DSERR_NOAGGREGATION"
                '    Description = "The object does not support aggregation."
                Case DSERR_NODRIVER
                    ErrorID = "DSERR_NODRIVER"
                    Description = "No sound driver is available for use."
                Case DSERR_NOINTERFACE
                    ErrorID = "DSERR_NOINTERFACE"
                    Description = "The requested COM interface is not available."
                Case DSERR_OTHERAPPHASPRIO
                    ErrorID = "DSERR_OTHERAPPHASPRIO"
                    Description = "Another application has a higher priority level, preventing this call from succeeding."
                Case DSERR_OUTOFMEMORY
                    ErrorID = "DSERR_OUTOFMEMORY"
                    Description = "The DirectSound subsystem could not allocate sufficient memory to complete the caller's request."
                Case DSERR_PRIOLEVELNEEDED
                    ErrorID = "DSERR_PRIOLEVELNEEDED"
                    Description = "The caller does not have the priority level required for the function to succeed."
                Case DSERR_UNINITIALIZED
                    ErrorID = "DSERR_UNINITIALIZED"
                    Description = "The DirectSound device has not been initialized."
                Case DSERR_UNSUPPORTED
                    ErrorID = "DSERR_UNSUPPORTED"
                    Description = "The function called is not supported at this time."
            End Select
    End Select
    
    If ErrorID = "" Then
        ErrorID = "VB_Error"
        Description = Error$(TempErr)
    End If
    
End Sub
