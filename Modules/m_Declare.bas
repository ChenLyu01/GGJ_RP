Attribute VB_Name = "m_Declare"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Public Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long

 
Public SoftwareRegPSD As String
Public User_Code As String
Public User_Codeing As Boolean

'用户注册部分
Public System_Sm_02(4) As String
Public System_Sm_03(4) As String


 '聊天泡泡的最大宽度
Public Const BubbleMaxWidth As Long = 140
Public Const SpriteFont_MaxLineLen As Long = 600

Private Type VFH
    BitmapWidth As Long         '位图的大小
    BitmapHeight As Long
    CellWidth As Long           '每字
    CellHeight As Long
    BaseCharOffset As Byte      '开始字
    CharWidth(0 To 255) As Byte '字符实际宽度
End Type

Public Type CustomFont
    HeaderInfo As VFH           '标题
    RowPitch As Integer         '每行的字符数量
    RowFactor As Single         '每个字符占用的纹理宽度
    ColFactor As Single         '同上，高度
    CharHeight As Byte          '文本的高度
End Type


Type m_Position
    x As Integer
    y As Integer
    Enable As Boolean
End Type

Type m_RECT
    Top           As Integer
    bottom        As Integer
    Left          As Integer
    Right         As Integer
End Type

Type m_WandH
    Width As Integer
    Height As Integer
End Type

Type m_GraphicPosition
    x As Integer
    y As Integer
    Width As Integer
    Height As Integer
End Type

Type m_Npcplay      '可以删除的类型
    C_Position    As m_Position '当前位置
    G_Position    As m_Position '寻路中使用的目标位置
    Next_Position As m_Position '寻路中下一个位置位置
    Roading(100)  As m_Position    '则路
    RoadData      As Integer '则路指针
    RodeTry       As Boolean
    RoadXY        As Boolean
    G_Event(1024) As Byte  '目标位置
    MoveDirection As Byte
    Alive         As Boolean '转世
    SleepTimer    As Integer '停滞时间
    EventTimer    As Integer '事件计数器
    EventBreak    As Boolean '事件激活器
    PositionTimer As Single '位置计时器
    MoveTimer     As Integer '人物动画计时器
    Agi           As Integer '速度
    Pic           As Integer
    MoveSpeed     As Single  '移动速度
    Hp            As Integer '生命力
    MaxHp         As Integer '人物最大血量
    Def           As Integer '防御力
    Str           As Integer '力量
    Int           As Integer '智力
    Men           As Integer '精神
    Att           As Integer '攻击力
    AttSpeed      As Integer '攻击速度
    level         As Integer '等级
    Exp           As Long    '经验
    Mp            As Integer '魔法量
    MaxMp         As Integer '最大魔法量
    Width         As Integer '宽
    Height        As Integer '高
    HeadPic       As Integer '头像图
    MagicBall     As Boolean
    MagicBallTimer As Integer
End Type

Type m_Clock
    myHour As Byte
    myMinute As Byte
End Type

Type m_MapGraphicPosition
    GraphicID As Byte
    GraphicPosition As m_GraphicPosition
    MapPosition(32) As m_Position
    LayerNum As Byte
    GraphicName As String
End Type

Type m_Buffer
    BackBuffer As Long
    BackBufferBmp As Long
    OldBackBufferDC As Long
    OldTilesetBmpDC(16) As Long
    TileSetBmp(16) As Long
End Type

Type m_Windows
    GraphicID As Byte
    Enable As Boolean
    Visible As Boolean
    LoadPosition As m_GraphicPosition
    DrawPosition As m_Position
End Type

Type m_BlockTile
    GraphicID As Byte
    GraphicPosition As m_GraphicPosition
End Type


Type m_Event
    m_Name As String
    m_Description As String
    PicPosition As m_Position
End Type

Type m_Effect
    GraphicID As Byte
    Visible(2) As Boolean
    FrameRun(2) As Boolean
    Timer(2) As Integer
    FrameCount As Byte
    Matrix As m_Position
    MatrixGraphic As m_WandH
    DrawPosition(2) As m_Position
    LoadPosition As m_GraphicPosition
End Type

Type m_steps
    GraphicID As Byte
    Arrow(4) As m_GraphicPosition
    Direction As Integer
    Position As m_Position
    Visible As Boolean
End Type

Type m_Map
    sName As String
    ID As Integer
    Tiles_Posi(3) As m_Position
    TileID As Byte
    TilesInfo As m_GraphicPosition
    Tile_Map As m_WandH
    Tile_Object As m_WandH
    Tiles(5) As m_MapGraphicPosition
    Letters(4, 26) As m_MapGraphicPosition
    Letter(52) As m_MapGraphicPosition
    BlockTilesInfo(16) As m_BlockTile
    BlockTilesID As Byte
    Objects(32) As m_MapGraphicPosition
    Events(16, 14) As Byte
    Block(16, 14) As Byte
    BlockTiles(16, 14) As Byte
    GameEvents(64) As m_Event
    Windows(4) As m_Windows
    Effect(8) As m_Effect
End Type

Type m_GraphicFiles
    FilesName As String
    BlackPicPixel As m_GraphicPosition
End Type
        
Type m_Player
    GraphicID As Byte
    Tile As m_WandH
    Pic As m_Position
    Pic_Temp(2) As m_Position
    Info As m_Npcplay
    mName As String
End Type

Public Type m_Font
    x As Integer  '字体X坐标
    y As Integer   '字体Y坐标
    G_String As String '字体内容
    G_MaxLineLen As Integer '字体内容宽度
    G_WordWraped As Boolean
    Visiable As Boolean
End Type

Public Type m_Story
    m_Name As String
    Text As String
End Type

Type m_Code
    mName As String
    Text As String
    Autorun As Boolean
    runFlag As Boolean
    Order As String
End Type

Type m_Graphic
    EffectTimer(8) As Single
    Posi  As m_Position
    RECT As m_RECT
    Screen As m_WandH
    Buffer As m_Buffer
    Map As m_Map
    GraphicFiles(16) As m_GraphicFiles
    Player(2) As m_Player
    SpriteFont(100) As m_Font
    WindowsFont(100) As m_Font
    Font_Default As CustomFont   '默认字体
    Clock As m_Clock
    Code As m_Code
    AILoaded As Boolean
End Type

Type m_FilePath
    Graphics As String
    Code As String
    CourseMap As String
    CodeName As String
    Story As String
End Type

Type m_Switch
    Object As Boolean
    Block As Boolean
    Event As Boolean
    Player As Boolean
    Effect As Boolean
    Pathway As Boolean
    Debug As Boolean
    Letters As Boolean
    Timer As Boolean
    Steps As Boolean
    PathwayWithBlock As Boolean
    Talk As Boolean
End Type



Public this_AI_text As m_Position

Public Steps(100) As m_steps

 
 

Public this_Graphic As m_Graphic
Public this_FilePath As m_FilePath
Public this_Switch As m_Switch

Public Const MoveLeft = 1
Public Const MoveRight = 2
Public Const MoveUp = 3
Public Const MoveDown = 0
Public CurrentPlayerID As Byte
Public CurrentSteps As Byte
Public CurrentEventTimer As Integer
Public CurrentGameEvents As Byte
Public EventBreak As Boolean
