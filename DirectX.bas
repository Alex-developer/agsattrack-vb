Attribute VB_Name = "DirectXStuff"
'Option Explicit
'
'Public Const VIEWPORT_WIDTH = 1024
'Public Const VIEWPORT_HEIGHT = 512
'Public Const PI = 3.14159265358979
'
'Public D3DRM As IDirect3DRM
'
''***********************************
''*  Function to create a Sphere **
''***********************************
'
'Public Sub BuildSphere(objMeshBuilder As IDirect3DRMMeshBuilder)
'    Dim aVertices(1 To 1000) As D3DVECTOR
'    Dim aNormals(0) As D3DVECTOR
'    Dim aFaces(1 To 10000) As Long
'    Dim intVertices As Long
'    Const STEPA = 10
'    Const STEPB = 10
'    Dim axeZ As D3DVECTOR, origine As D3DVECTOR, AxeY As D3DVECTOR
'    origine.x = 0:   origine.y = 1:   origine.z = 0
'    axeZ.x = 0:      axeZ.y = 0:      axeZ.z = 1
'    AxeY.x = 0:      AxeY.y = 1:      AxeY.z = 0
'    intVertices = 1
'    Dim i As Integer, j As Integer
'    Dim tmp As D3DVECTOR
'    For i = STEPA To 180 - STEPA Step STEPA
'    For j = 0 To 360 - STEPB Step STEPB
'            D3DRMVectorRotate tmp, origine, axeZ, i * PI / 180
'            D3DRMVectorRotate aVertices(intVertices), tmp, AxeY, j * PI / 180
'            intVertices = intVertices + 1
'       Next
'    Next
'    intVertices = intVertices - 1
'    Dim Index As Integer
'    Index = 1
'    For i = STEPA To 180 - 2 * STEPA Step STEPA
'        Dim FirstIndex As Long
'        FirstIndex = Index
'        For j = 0 To 360 - STEPB Step STEPB
'            aFaces(Index) = 4
'            aFaces(Index + 1) = (Index \ 5) + 1
'            aFaces(Index + 2) = (Index \ 5)
'            aFaces(Index + 3) = ((Index \ 5) + (360 \ STEPB))
'            aFaces(Index + 4) = (Index \ 5) + 1 + (360 \ STEPB)
'            If j = 360 - STEPB Then
'                aFaces(Index + 1) = FirstIndex \ 5  '+ 1
'                aFaces(Index + 4) = FirstIndex \ 5 + (360 \ STEPB)
'            End If
'            Index = Index + 5
'        Next
'    Next
'    aFaces(Index) = (360 / STEPB) - 1
'    Index = Index + 1
'    For i = 1 To (360 / STEPB) - 1
'        aFaces(Index) = i
'        Index = Index + 1
'    Next
'    aFaces(Index) = 360 / STEPB
'    Index = Index + 1
'    For i = 0 To (360 / STEPB) - 1
'        aFaces(Index) = intVertices - i - 1
'        Index = Index + 1
'    Next
'    aFaces(Index) = 0
'    objMeshBuilder.AddFaces intVertices, aVertices(1), 0, aNormals(0), aFaces(1), Nothing
'  End Sub
'Public Sub PutSphereTexture(D3DRM As IDirect3DRM, MeshBuilder As IDirect3DRMMeshBuilder, ByVal strTextureFileName As String)
'    Dim Box As D3DRMBOX
'    Dim MaxY As Single, MinY As Single
'    Dim Height As Single
'    Dim Wrap As IDirect3DRMWrap
'    Dim Texture As IDirect3DRMTexture
'    ' Bounding box
'    MeshBuilder.GetBox Box
'    MaxY = Box.Max.y
'    MinY = Box.Min.y
'    Height = MaxY - MinY
'    D3DRM.CreateWrap D3DRMWRAP_CYLINDER, Nothing, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, MinY / Height, 1, 1 / Height, Wrap
'    Wrap.Apply MeshBuilder
'    D3DRM.LoadTexture strTextureFileName, Texture
'    MeshBuilder.SetTexture Texture
'End Sub
