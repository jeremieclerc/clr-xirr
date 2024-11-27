IF (SELECT TOP 1 value FROM sys.configurations WHERE name = 'clr enabled') != 1
BEGIN
    EXEC sp_configure 'clr enabled', 1
    RECONFIGURE
END
DECLARE @sql nvarchar(MAX), @hash varbinary(64)
DECLARE @clrBinary varbinary(MAX) = 0x4D5A90000300000004000000FFFF0000B800000000000000400000000000000000000000000000000000000000000000000000000000000000000000800000000E1FBA0E00B409CD21B8014CCD21546869732070726F6772616D2063616E6E6F742062652072756E20696E20444F53206D6F64652E0D0D0A2400000000000000504500004C0103005AFAFF660000000000000000E00022200B0130000014000000060000000000009E3200000020000000400000000000100020000000020000040000000000000006000000000000000080000000020000000000000300608500001000001000000000100000100000000000001000000000000000000000004C3200004F000000004000009802000000000000000000000000000000000000006000000C000000143100001C0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000200000080000000000000000000000082000004800000000000000000000002E74657874000000A4120000002000000014000000020000000000000000000000000000200000602E7273726300000098020000004000000004000000160000000000000000000000000000400000402E72656C6F6300000C0000000060000000020000001A000000000000000000000000000040000042000000000000000000000000000000008032000000000000480000000200050054280000C00800000100000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000360002730500000A7D010000042A000013300400CF00000001000011000F02280600000A2D2E0F01280700000A2D250F01230000000000000000280800000A8C0A000001FE160A0000016F0900000A16FE012B01160A06398E0000000004280A00000A280B00000A280C00000A0C1202280D00000A0B027B01000004078C120000016F0E00000A15FE010D092C20027B01000004078C120000010F01280F00000A8C130000016F1000000A002B3B00027B01000004078C120000016F1100000AA5130000011304027B01000004078C1200000111040F01280F00000A588C130000016F1200000A0000002A0013300500B90000000200001100037B010000040A066F1300000A0B066F1400000A0C160D388900000000027B0100000407096F1500000A6F0E00000A15FE01130411042C26027B0100000407096F1500000A08096F1500000AA5130000018C130000016F1000000A002B4200027B0100000407096F1500000A6F1100000AA5130000011305027B0100000407096F1500000A110508096F1500000AA513000001588C130000016F1200000A0000000917580D09066F1600000AFE04130611063A65FFFFFF2A00000013300300D50000000300001100027B010000046F1300000A0A027B010000046F1400000A0B170C170D1613042B48000711046F1500000AA513000001230000000000000000FE02130611062C02160C0711046F1500000AA513000001230000000000000000FE04130711072C02160D110417581304001104076F1700000A2F050809602B0116130811082DA2080960130911092C0A007E1800000A130A2B3F00020728070000067D020000040706027B02000004280800000613051105281900000A130B110B2C0A007E1800000A130A2B0C001105731A00000A130A2B00110A2A0000001330030054000000040000110002036F1B00000A7D0200000402730500000A7D01000004036F1C00000A0A160B2B2800027B01000004036F1C00000A8C12000001036F1B00000A8C130000016F1000000A00000717580B0706FE040C082DD02A133003007B000000050000110003027B020000046F1D00000A0003027B010000046F1600000A6F1E00000A00027B010000046F1300000A0A027B010000046F1400000A0B160C2B2C000306086F1500000AA5120000016F1E00000A000307086F1500000AA5130000016F1D00000A00000817580C08027B010000046F1600000AFE040D092DC22A00133003006100000006000011002300000000000000000A160B2B15000602076F1500000AA513000001580A000717580B07026F1700000AFE040C082DDE06230000000000000000FE0516FE010D092C0D239A9999999999B93F13042B0D239A9999999999B9BF13042B0011042A00000013300500760000000700001100040A238DEDB5A0F7C6B03E0B1F640C020304070828090000060A06281900000A2D0E0623000000000000F0BFFE022B01160D092C050613042B38020304280A0000060A06281900000A2D0E0623000000000000F0BFFE022B0116130511052C050613042B0D23000000000000F8FF13042B0011042A0000133004004D0100000800001100040A160D3825010000002300000000000000000B2300000000000000000C16130438BA000000000311046F1500000AA51200000103166F1500000AA512000001596C230000000000D076405B13050623000000000000F03F5813061106230000000000000000FE0316FE01130811082C110023000000000000F8FF130938C700000011061105281F00000A13071107230000000000000000FE01130A110A2C0D0023BBBDD7D9DF7CDB3D130700070211046F1500000AA51300000111075B580B0811050211046F1500000AA5130000015A110711065A5B590C001104175813041104026F1700000AFE04130B110B3A33FFFFFF07282000000A05FE04130C110C2C06000613092B4108282000000A05FE04130D110D2C0C0023BBBDD7D9DF7CDB3D0C000607085B590A0917580D00090E04FE04130E110E3ACDFEFFFF23000000000000F8FF13092B0011092A000000133003003A01000009000011002300000000000000000A23F168E388B5F8E43E0B040C230000000000005940130423EA8CA039593E2946130523000000000000F03F130623AE47E17A14AEEF3F130723000000000000F03F13081F64130916130A080D2B6F0011071308080D080203280B00000613071108110759282000000A13041107230000000000000000FE0516FE01130B110B2C07081106580C2B150011062300000000000000405B1306081106590C00110A1758130A11041105FE02130C110C2C0D23000000000000F8FF130D2B70001104073608110A1109FE042B0116130E110E3A7AFFFFFF090A120023000000000000F07F282100000A2D12120023000000000000F0FF282100000A2B0117130F110F2C0D23000000000000F8FF130D2B1E110A1109FE01131011102C0D23000000000000F8FF130D2B0506130D2B00110D2A000013300600DE0000000A000011002300000000000000000A160B38B20000000023000000000000F03F02586C04076F1500000AA51200000104166F1500000AA512000001596C230000000000D076405B6C281F00000A0D1203230000000000000000282100000A0C082C1B0603076F1500000AA513000001239A9999999999B93F5B580A2B460603076F1500000AA51300000123000000000000F03F02586C04076F1500000AA51200000104166F1500000AA512000001596C230000000000D076405B6C281F00000A5B580A000717580B07036F1700000AFE04130411043A3CFFFFFF0613052B0011052A000042534A4201000100000000000C00000076342E302E33303331390000000005006C00000058030000237E0000C40300005803000023537472696E6773000000001C070000040000002355530020070000100000002347554944000000300700009001000023426C6F620000000000000002000001571702000902000000FA013300160000010000001500000003000000020000000B000000140000000100000021000000040000000A000000010000000200000001000000000038010100000000000600DA0023020600FA0023020600A50010020F00430200000A00C602DC010A00B900DC0106007B0062010A002201DC010600010389020A00650052020A006F0052020600690162010600FB0289020600C2011A000600CF011A000600CD026201060072006201060005006201060068006201060081018902060033016201000000000B00000000000100010009211000240000001D000100010082011000F70100004100030007000600AC02FA000600A602FE005020000000008600EC020600010060200000000086008500010101003C21000000008600540009010300042200000000860090000F010400E82200000000E6014B0014010400482300000000E6019F001A010500D0230000000096009C02200106004024000000009600290026010700C4240000000091008D012F010A002026000000009600320326010F006827000000009100AC013A011200000001006D0200000200670200000100BC01000001000A02000001002603000001006D02000001006D0200000200C10200000300A602000001006D0200000200C10200000300A602000004004C01000005007B02000001006D0200000200C10200000300A602000001009A00000002006D0200000300C10202002100090006020100110006020600190006020A00310006021000490006020600590041011F00510041011F005100D40223008100740229005900E0022E00890028033500890072013A006100B802430049004A034700510018014C0049005000500049005001560049005901500049001903680049000C036800690050016D004900F1024300A100F1024300510047018400990014008800510006028D0071005A004C0071000100430079009F008D0079009F000100A9002403C400A9000C02CA0099007402E3002E000B0043012E0013004C012E001B006B0143002300740116005B00720092009800A100A900B200CF00E80004800000000000000000000000000000000024000000040000000000000000000000F100420000000000040000000000000000000000F10036000000000003000200000000000052656164496E743332003C4D6F64756C653E0049734E614E0053797374656D2E494F00584952520043616C63756C6174654952520053797374656D2E44617461006D73636F726C6962005265616400416464004D657267650052656164446F75626C650053716C446F75626C650053716C4461746554696D650056616C75655479706500416363756D756C617465005465726D696E61746500726174650057726974650044656275676761626C654174747269627574650053716C55736572446566696E656441676772656761746541747472696275746500436F6D70696C6174696F6E52656C61786174696F6E734174747269627574650052756E74696D65436F6D7061746962696C697479417474726962757465006765745F56616C7565004942696E61727953657269616C697A65004D61746800584952522E646C6C006765745F49734E756C6C00746F6C006765745F4974656D007365745F4974656D0053797374656D0054696D655370616E006F705F5375627472616374696F6E0049436F6C6C656374696F6E0043616C63756C6174654952525573696E674E6577746F6E52617068736F6E0043616C63756C61746552657475726E0047726F75700042696E6172795265616465720042696E617279577269746572004D6963726F736F66742E53716C5365727665722E536572766572005869727243616C63756C61746F72002E63746F72004162730053797374656D2E446961676E6F73746963730053797374656D2E52756E74696D652E436F6D70696C6572536572766963657300446562756767696E674D6F6465730053797374656D2E446174612E53716C54797065730064617465730076616C75657300457175616C73006D6178497465726174696F6E730053797374656D2E436F6C6C656374696F6E73004175746F477565737300677565737300697272456C656D656E7473006765745F44617973006461797300466F726D6174004F626A656374006F705F496D706C69636974006F705F4578706C6963697400496E6974006765745F436F756E7400494C69737400536F727465644C6973740047657456616C75654C697374004765744B65794C69737400506F77006765745F546F6461790043616C63756C6174654952525573696E674C656761637900496E6465784F664B657900000000000000001605AD18948A7C4083E31B35DB457AF80004200101080320000105200101111105200101111508070502081131020D0320000205000111290D042001021C0600011145112D040000114508000211311145114503200008042001081C0320000D052002011C1C0420011C1C0C070712251235123508020D0204200012350420011C0811070C123512350202080D0202020211290203061129040001020D042001010D0507030808020807041235123508020707050D0802020D0807060D0D08020D0211070F0D0D0D08080D0D0D020D02020202020500020D0D0D0400010D0D1307110D0D0D0D0D0D0D0D0D080802020D020202042001020D0807060D08020D020D08B77A5C561934E0890306122502060D072002011129112D052001011108042000112905200101123905200101123D0500010D12350800030D123512350D0A00050D123512350D0D080800030D0D123512350801000800000000001E01000100540216577261704E6F6E457863657074696F6E5468726F7773010801000701000000001A010002000000010054080B4D61784279746553697A65FFFFFFFF00000000005AFAFF6600000000020000001C01000030310000301300005253445342BE15466C70A7419AF85ADAAC6CCF3101000000433A5C55736572735C6D65796E69656C615C736F757263655C7265706F735C636C722D786972725C584952525C6F626A5C44656275675C584952522E70646200000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000007432000000000000000000008E32000000200000000000000000000000000000000000000000000080320000000000000000000000005F436F72446C6C4D61696E006D73636F7265652E646C6C0000000000FF2500200010000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000001001000000018000080000000000000000000000000000001000100000030000080000000000000000000000000000001000000000048000000584000003C02000000000000000000003C0234000000560053005F00560045005200530049004F004E005F0049004E0046004F0000000000BD04EFFE00000100000000000000000000000000000000003F000000000000000400000002000000000000000000000000000000440000000100560061007200460069006C00650049006E0066006F00000000002400040000005400720061006E0073006C006100740069006F006E00000000000000B0049C010000010053007400720069006E006700460069006C00650049006E0066006F0000007801000001003000300030003000300034006200300000002C0002000100460069006C0065004400650073006300720069007000740069006F006E000000000020000000300008000100460069006C006500560065007200730069006F006E000000000030002E0030002E0030002E003000000032000900010049006E007400650072006E0061006C004E0061006D006500000058004900520052002E0064006C006C00000000002800020001004C006500670061006C0043006F0070007900720069006700680074000000200000003A00090001004F0072006900670069006E0061006C00460069006C0065006E0061006D006500000058004900520052002E0064006C006C0000000000340008000100500072006F006400750063007400560065007200730069006F006E00000030002E0030002E0030002E003000000038000800010041007300730065006D0062006C0079002000560065007200730069006F006E00000030002E0030002E0030002E003000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000003000000C000000A03200000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000 

SET @hash = (SELECT TOP 1 hash FROM sys.trusted_assemblies WHERE description = 'AssemblyXirr')
IF @hash IS NOT NULL EXEC sp_drop_trusted_assembly @hash
DROP AGGREGATE IF EXISTS dbo.XIRR
DROP ASSEMBLY IF EXISTS AssemblyXirr

SET @hash = HASHBYTES('SHA2_512', @clrBinary)
EXEC sp_add_trusted_assembly @hash, N'AssemblyXirr'
CREATE ASSEMBLY AssemblyXirr AUTHORIZATION [dbo] FROM @clrBinary WITH PERMISSION_SET = UNSAFE;

EXEC ('CREATE AGGREGATE [dbo].[XIRR] (@values [float], @dates [datetime]) RETURNS[float] EXTERNAL NAME [AssemblyXirr].[XIRR]')
