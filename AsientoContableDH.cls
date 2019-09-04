VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsientoContableDH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim aSubcuentas(5, 1000) As String
Dim msModulo As String
Dim msDebHab As String
Dim msMainTb As String 'Alias de tabla de la registración
Dim msHeadTb As String 'Tabla encabezado, en paso de ser un pago o una modificación difiere del Alias de la registración
Dim msDebTb As String 'Tabla Conceptos Debe
Dim msDebTbD As String 'Tabla Conceptos Debe Dimensiones
Dim msDebTbG As String 'Tabla Conceptos Debe Subcuenta
Dim msHabTb As String 'Tabla Conceptos Haber
Dim msHabTbD As String 'Tabla Conceptos Haber Dimensiones
Dim msHabTbG As String 'Tabla Conceptos Haber Subcuenta
Dim sTipComReg As String 'Según si en el módulo el campo se trata como tipreg o tipcom
Dim sCampoActivo As String
Dim mvEvent As Integer
Dim msAperturaCJ As Integer 'Marca si ya fue grabada la apertura contable de CJ
Dim mPasaValImport As Boolean
Dim bEsAnuTot As Boolean, bNoRecuperaAsientoContable As Boolean
Dim bControladoCircuito As Boolean
Dim bReady As Boolean

Const cCantidad As Integer = 0
Const cCodiRL As Integer = 1
Const cCodiRP As Integer = 2
Const cSubCuenta As Integer = 3
Const cImporte As Integer = 4
Const cCodVin As Integer = 5
Const cDesSubCue As Integer = 6
Const cDesRL As Integer = 7
Const cDesRP As Integer = 8
Const cDesVin As Integer = 9


Public Function ThisEvent(Who As Object, vEvent As Variant) As Integer
    ThisEvent = 0
    mvEvent = vEvent
    Select Case CInt(vEvent)
    Case 13 'READY NEW
        DefinoModulo Who
        DefinoTablas Who
        Valida_Import Who 'Agregado por BFM 20190729 para que no de error al iniciar una registración que trae datos completos como la Boleta de depósito
        bReady = True
        
    Case 23 'VALIDATE FIELD
        'Esto lo hago porque en ciertos links que vienen de un paso anterior, el primer evento es VALIDATE_FIELD
        If msModulo = "" Then
            DefinoModulo Who
            DefinoTablas Who
        End If
            
        ThisEvent = ValidoCampos(Who)
    'Case 24 'READY EDIT
    Case 27 'SET ACTIVE TABLE
        ThisEvent = ValidoTablas(Who)
        If Who.CurrentTable = msHabTbD Then
            ThisEvent = BuscaAperturaCtble(Who, msHabTb, msHabTbD, msHabTbG)
            Who.IntoEvent = True
        End If
        If Who.CurrentTable = msDebTbD Then
            ThisEvent = BuscaAperturaCtble(Who, msDebTb, msDebTbD, msDebTbG)
            Who.IntoEvent = True
        End If
        
    Case 28 'PRE VALIDATE SAVE
    ThisEvent = BuscaAperturaCtble(Who, msHabTb, msHabTbD, msHabTbG)
    ThisEvent = BuscaAperturaCtble(Who, msDebTb, msDebTbD, msDebTbG)
    
    'BFM 20190625 quitado el control con la condición 1 = 2, hay que probar el funcionamiento (causaba problemas en todos los módulos)
    If msHeadTb <> "CJRMVH" And msHeadTb <> "FCRMVH" And 1 = 2 Then
    
        Dim oRowI_
        Dim oRowD_
        For Each oRowI_ In Who.Table.Rows(1).Tables(msDebTb).Rows
            oRowI_.Tables(msDebTbD).Rows.MaxRows = oRowI_.Tables(msDebTbD).Rows.MaxRows + 1
            'On Error Resume Next
            oRowI_.Tables(msDebTbD).Rows.Add (2)
            oRowI_.Tables(msDebTbD).Rows.Remove (2)
            'On Error GoTo 0
            oRowI_.Tables(msDebTbD).Rows.MaxRows = oRowI_.Tables(msDebTbD).Rows.MaxRows - 1
        Next
    End If
    ThisEvent = ValidaFecha(Who)
    
    Case 38 'POST SAVE
    Grabo_RL_RP_VIN Who, msDebTb, msDebTbD, msDebTbG
    Grabo_RL_RP_VIN Who, msHabTb, msHabTbD, msHabTbG
    
    End Select
            
End Function

Private Sub DefineSiEsAnulacionTotal(Who As Object)

    Dim sSql As String
    Dim oRs As Object
    Dim oField As Object
    Dim bExisteCircuito As Boolean
    
    bControladoCircuito = True
    
    With Who.Table.Rows(1)
    
        bExisteCircuito = ExisteCampo(Who.Table.Rows(1).Fields, msModulo & "RMVH_CIRCOM")
        
        If bExisteCircuito Then
        
            sSql = "SELECT " & msModulo & "TCIH_ANULAC, " & msModulo & "TCIH_ANUTOT, " & msModulo & "TCIH_NTRACN"
            sSql = sSql & " FROM " & msModulo & "TCIH"
            sSql = sSql & " WHERE " & msModulo & "TCIH_CIRCOM = " & .Fields(msModulo & "RMVH_CIRCOM").SQLValue
            sSql = sSql & " AND " & msModulo & "TCIH_CIRAPL = " & .Fields(msModulo & "RMVH_CIRAPL").SQLValue
    
            Set oRs = Who.DataAccess.OpenResultset(sSql)
            
            If Not oRs.EOF Then
        
                bEsAnuTot = (oRs(msModulo & "TCIH_ANULAC").Value = "S" And oRs(msModulo & "TCIH_ANUTOT").Value = "S")
                bNoRecuperaAsientoContable = oRs(msModulo & "TCIH_NTRACN").Value = "S"
        
            End If
        
            oRs.Close

        End If

    End With
    


End Sub




Private Function ValidaFecha(Who As Object) As Integer
Dim sSqlVF As String, RsVF

    sSqlVF = "EXEC SP_ValidaFecha "
    sSqlVF = sSqlVF & " @MODFOR = '" & Who.Table.Rows(1).Fields(msHeadTb & "_MODFOR").Value & "', "
    sSqlVF = sSqlVF & " @CODFOR = '" & Who.Table.Rows(1).Fields(msHeadTb & "_CODFOR").Value & "', "
    sSqlVF = sSqlVF & " @FCHMOV = '" & Who.Table.Rows(1).Fields(msHeadTb & "_FCHMOV").Value & "'"
    
    Set RsVF = Who.DataAccess.OpenResultset(CStr(sSqlVF))

    If Not RsVF.EOF Then
        If RsVF("Valor").Value = "S" Then
            Who.SetError "GR000068", "La fecha corresponde a un periodo Cerrado"
            ValidaFecha = 1   ' Cantidad maxima de apertura
        End If
    End If
    
    RsVF.Close
    Set RsVF = Nothing
    
End Function
'Esta funcion se usa en la interfaz de CJ
Public Sub AsientosDH(Who As Object, oWizard As Object)
    Dim Rs, sSql, sSQL2, Rs2
    Dim oRowI01 As Object, oRowD01 As Object, oRowG01 As Object
    Dim oRowI02 As Object, oRowD02 As Object, oRowG02 As Object
    Dim nIndice As Integer, nIndice2 As Integer, nNroLot As Integer, nNroIlo As Integer
    Dim sModCom, sCodCom, sCirCom, sCirApl, sTipReg, sNroCta, sTipCpt, sCodCpt, sCuenta, sTipPro, sArtCod, sDeposi, sSector, sDepIng, sSecIng, sNotros, sCodigoSN, sDescSN, sModGen, sComGen
    Dim nItmSub As Integer, nImporte As Double
    Dim nCtaMax, sCtaISC, sCtaEdi, nSCsMax, sSCIEdi, sRLsVal As String, iRLsItm, sRPsVal As String, sRLsEdi, sRPsEdi, sCodVin, sCVsEdi, sCVsReq, sCVsCnd
    Dim sNroCVT, sNroCPV, sNroCVTCPV As String
    Dim oRowGx As Object
    
    If Who.Table.Name = "" Or Who.Table.Name = "CJRMVH" Then
            msMainTb = "CJRMVH"
            msHeadTb = "CJRMVH"
            msDebTb = "CJRMVI01"
            msHabTb = "CJRMVI02"
            msDebTbD = "CJRMVD01"
            msHabTbD = "CJRMVD02"
            msDebTbG = "CJRMVG01"
            msHabTbG = "CJRMVG02"
            sTipComReg = "TIPREG"
    End If
    
    
    For Each oRowI01 In Who.Table.Rows(1).Tables(msDebTb).Rows
        
        Who.Table.Rows(1).Tables(msDebTb).Activate
        
        If msDebHab = "" Then
            If msHeadTb = "FCRMVH" Or msHeadTb = "VTRMVH" Then
                AsignoValor Who, oRowI01.Fields("CJRMVI_NROCVT"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
                AsignoValor Who, oRowI01.Fields("CJRMVI_NROSVT"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
            End If
            If msHeadTb = "CORMVH" Or msHeadTb = "PVRMVH" Then
                AsignoValor Who, oRowI01.Fields("CJRMVI_NROCPV"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
                AsignoValor Who, oRowI01.Fields("CJRMVI_NROSPV"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
            End If
        End If

        sNroCVT = ""
        sNroCPV = ""
        
        If oRowI01.Fields("CJRMVI_IMPORT").Value = 0 Then
            sSql = " Insert into INTERR (INTERR_NROITM, INTERR_CODTRA, INTERR_FCHMOV, INTERR_FIELDS, INTERR_VALUES, "
            sSql = sSql & " INTERR_ERRMSG , INTERR_FECALT, INTERR_FECMOD, INTERR_HORMOV, INTERR_ULTOPR, INTERR_DEBAJA, "
            sSql = sSql & " INTERR_USERID,INTERR_OALIAS)"
            sSql = sSql & "  Values ((SELECT MAX(IN2.INTERR_NROITM)+1 FROM INTERR AS IN2), 'INTCJM', GETDATE(),"
            sSql = sSql & " 'CJRMVH_CODFOR|USR_CJRMVH_TIPREG|CJRMVH_FCHMOV|CJRMVH_NROCVT|CJRMVH_NROCPV|', "
            sSql = sSql & "'" & Who.Table.Rows(1).Fields(msHeadTb & "_CODFOR").Value & "|"
            sSql = sSql & Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_" & sTipComReg).Value & "|"
            sSql = sSql & CStr(Year(Date)) & Right("00" & CStr(Month(Date)), 2) & Right("00" & CStr(Day(Date)), 2) & "|||',"
            sSql = sSql & "'En el registro " & oRowI01.Fields("CJRMVI_NROITM").Value & " : Imposible grabar con Importe en 0'"
            sSql = sSql & ",  GETDATE(),  GETDATE(),  NULL , 'A', 'N', 'DLL', 'INTERR')"
            Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
            Rs.Close
            Exit Sub
        End If
    
        nImporte = oRowI01.Fields("CJRMVI_IMPORT").Value
        sModCom = "CJ"
        sCodCom = oRowI01.Fields("CJRMVI_CODFOR").Value
        sCirCom = ""
        sCirApl = ""
        sTipReg = Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_" & sTipComReg).Value
        sTipCpt = oRowI01.Fields("CJRMVI_TIPCPT").Value
        sCodCpt = oRowI01.Fields("CJRMVI_CODCPT").Value
        sCuenta = oRowI01.Fields("CJRMVI_CUENTA").Value
        sTipPro = ""
        sArtCod = ""
        sDeposi = ""
        sSector = ""
        sDepIng = ""
        sSecIng = ""
        sNotros = ""
        nNroLot = oRowI01.Fields("USR_CJRMVI_NROLOT").Value
        nNroIlo = oRowI01.Fields("USR_CJRMVI_NROILO").Value
        sModGen = oRowI01.Fields("USR_CJRMVI_MODGEN").Value
        sComGen = oRowI01.Fields("USR_CJRMVI_COMGEN").Value
        
        sSql = "SELECT ISNULL(USR_CJIRVI_TIPPRO,'') sTipProIN, ISNULL(USR_CJIRVI_ARTCOD,'') sArtCodIN, ISNULL(USR_CJIRVI_CODVIN,'') sNotrosIN,"
        sSql = sSql & " USR_CJIRVI_NROCVT sNroCVTIN,USR_CJIRVI_NROCPV sNroCPVIN "
        sSql = sSql & " FROM USR_CJIRVI "
        sSql = sSql & "  Where USR_CJIRVI_NROLOT = " & nNroLot
        sSql = sSql & " And USR_CJIRVI_MODGEN = '" & sModGen & "'"
        sSql = sSql & " And USR_CJIRVI_COMGEN = '" & sComGen & "'"
        sSql = sSql & " And USR_CJIRVI_NROILO = " & nNroIlo
        Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
        If Not Rs.EOF Then
            If Trim(Rs("sTipProIN").Value) <> "" Then
               sTipPro = Trim(Rs("sTipProIN").Value)
            End If
            If Trim(Rs("sArtCodIN").Value) <> "" Then
               sArtCod = Trim(Rs("sArtCodIN").Value)
            End If
            If Trim(Rs("sNotrosIN").Value) <> "" Then
               sNotros = Trim(Rs("sNotrosIN").Value)
            End If
            If Trim(Rs("sNroCVTIN").Value) <> "" Then
               sNroCVT = Trim(Rs("sNroCVTIN").Value)
               oRowI01.Fields("CJRMVI_NROCVT").Value = sNroCVT
            End If
            If Trim(Rs("sNroCPVIN").Value) <> "" Then
               sNroCPV = Trim(Rs("sNroCPVIN").Value)
               oRowI01.Fields("CJRMVI_NROCPV").Value = sNroCPV
            End If
            
        End If
        Rs.Close
        
        If Trim(sNroCVT) <> "" Then
            sNroCVTCPV = Trim(sNroCVT)
        Else
           If Trim(sNroCPV) <> "" Then
                sNroCVTCPV = Trim(sNroCPV)
           End If
        End If
        sNroCta = sNroCVTCPV 'oRowI01.Fields("CJRMVI_NROCVT").Value
        
        sSql = "EXEC SP_CG_DLL_BuscoCTItem 'DLL', '" & sModCom & "','" & sCodCom & "','" & sCirCom & "','" & sCirApl & "','" & sTipReg & "','" & sNroCta & "','" & sTipCpt & "','" & sCodCpt & "','" & sCuenta & "','" & sTipPro & "','" & sArtCod & "','" & sDeposi & "','" & sSector & "','" & sDepIng & "','" & sSecIng & "', '" & sNotros & "'"
        Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
        If Not Rs.EOF Then
               nCtaMax = Rs("CuentaApertura").Value   ' Cantidad maxima de apertura
               sCuenta = Rs("CuentaValor").Value      ' Cuenta Contable
               sCtaISC = Rs("CuentaImpDim").Value     ' Cuenta es imputable in dimensiones
               sCtaEdi = Rs("CuentaEdita").Value      ' Cuenta es editable
               nSCsMax = Rs("SubCueApertura").Value   ' cantida maxima de apertura de subcuenta por cuenta
               sSCIEdi = "S"                          ' define si edito el importe
               sRLsVal = Rs("RLsValor").Value         ' Valor del RL
               iRLsItm = Rs("RLsItem").Value          ' Item de la USR_VTMRLS del Cliente
               sRPsVal = Rs("RPsValor").Value         ' Valor del RP
               sRLsEdi = Rs("RLsEdita").Value         ' RL es editable
               sRPsEdi = Rs("RPsEdita").Value         ' RP es editable
               sCodVin = ""                           ' Codigo vinulacion siempre es blanco
               sCVsEdi = Rs("CVsEdita").Value         ' CV es editable
               sCVsReq = Rs("CVsRequerido").Value     ' CV es Requerido
               sCVsCnd = Rs("CVsCondicion").Value     ' CV Condicion de validacion
        Else
            Exit Sub
        End If
        Rs.Close
        
        
        ' Me fijo si lo debo recuperar de la tabla de interfaces
'        If sRLsVal = "INTERF" Or sRPsVal = "INTERF" Then
'            sSQL = "SELECT USR_CJIRVI_CODIRL RLsValor, USR_CJIRVI_CODIRP RPsValor , isnull(USR_CJIRVI_CODVIN,'') sCodVin "
'            sSQL = sSQL & " FROM USR_CJIRVI "
'            sSQL = sSQL & "  Where USR_CJIRVI_NROLOT = " & nNroLot
'            sSQL = sSQL & " And USR_CJIRVI_MODGEN = '" & sModGen & "'"
'            sSQL = sSQL & " And USR_CJIRVI_COMGEN = '" & sComGen & "'"
'            sSQL = sSQL & " And USR_CJIRVI_NROILO = " & nNroIlo
'            Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSQL))
'            If Not Rs.EOF Then
'                If sRLsVal = "INTERF" Then
'                    sRLsVal = Rs("RLsValor").Value
'                End If
'                If sRPsVal = "INTERF" Then
'                    sRPsVal = Rs("RPsValor").Value
'                End If
'                If sRLsVal = "INTERF" And sRPsVal = "INTERF" Then
'                    sCodVin = Rs("sCodVin").Value
'                End If
'            End If
'            Rs.Close
'        End If


        sSql = "SELECT ISNULL(USR_CJIRVI_CODIRL,'') RLsValorIN, ISNULL(USR_CJIRVI_CODIRP,'') RPsValorIN  "
        sSql = sSql & " FROM USR_CJIRVI "
        sSql = sSql & "  Where USR_CJIRVI_NROLOT = " & nNroLot
        sSql = sSql & " And USR_CJIRVI_MODGEN = '" & sModGen & "'"
        sSql = sSql & " And USR_CJIRVI_COMGEN = '" & sComGen & "'"
        sSql = sSql & " And USR_CJIRVI_NROILO = " & nNroIlo
        Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
        If Not Rs.EOF Then
            If Trim(Rs("RLsValorIN").Value) <> "" And Trim(sRLsVal) = "" Then
               sRLsVal = Trim(Rs("RLsValorIN").Value)
            End If
            If Trim(Rs("RPsValorIN").Value) <> "" And Trim(sRPsVal) = "" Then
               sRPsVal = Trim(Rs("RPsValorIN").Value)
            End If
        End If
        Rs.Close

        ' aJUSTO LA i
        oRowI01.Fields("CJRMVI_CUENTA").ReadOnly = False
        If Trim(sCuenta) <> "" Then
            If oRowI01.Fields("CJRMVI_CUENTA").Value <> Trim(sCuenta) Then
               oRowI01.Fields("CJRMVI_CUENTA").Value = sCuenta
            End If
        Else
            sCuenta = oRowI01.Fields("CJRMVI_CUENTA").Value
        End If
        
        Erase aSubcuentas
        nIndice = 1
        If Left(sCuenta, 3) <> "CPT" And Left(sRLsVal, 3) <> "CPT" And Left(sRPsVal, 3) <> "CPT" Then
            'Me fijo y actualizo la cantidad de items
            aSubcuentas(0, 0) = nIndice
            
            'actualizo el array
            aSubcuentas(1, nIndice) = sRLsVal
            aSubcuentas(2, nIndice) = sRPsVal
            If sRLsVal <> "" And sRPsVal <> "" Then
                aSubcuentas(3, nIndice) = sRLsVal & sRPsVal
            Else
                aSubcuentas(3, nIndice) = ""
            End If
            aSubcuentas(4, nIndice) = nImporte
            aSubcuentas(5, nIndice) = ""
        
        Else
            If Left(sRLsVal, 3) = "CPT" Then
                RecuperoSCSolapa Who, "RL", Right(sRLsVal, 1), sRPsVal
            End If
            
            If Left(sRPsVal, 3) = "CPT" Then
                RecuperoSCSolapa Who, "RP", Right(sRPsVal, 1), sRLsVal
            End If
            
        End If
        
        
        
'       'Recorro la solapa Asiento Contable
'        'If oRowI01.Tables(msDebTbD).Rows.Count = 0 Then
'        With oRowI01.Tables(msDebTbD).Rows
'            oRowI01.Tables(msDebTbD).Rows.Add (oRowI01.Tables(msDebTbD).Rows.Count + 1)
'            Set oRowD01 = oRowI01.Tables(msDebTbD).Rows(oRowI01.Tables(msDebTbD).Rows.Count)
'            'With oRowD01
'        'End If
        For Each oRowD01 In oRowI01.Tables(msDebTbD).Rows
            oRowD01.SetError ""
            
            Dim BorroGD As Boolean
            BorroGD = True

            For Each oRowGx In oRowD01.Tables(msDebTbG).Rows
                If oRowGx.Fields("CJRMVG_SUBCUE").Value <> aSubcuentas(3, nIndice) And oRowGx.Fields("CJRMVG_SUBCUE").Value <> "" Then
                    BorroGD = False
                End If
            Next
            If (BorroGD) Then
                If oRowD01.Tables(msDebTbG).Rows.Count > 0 Then
                    oRowD01.Tables(msDebTbG).Rows.Clear
                End If
                For nIndice2 = 1 To aSubcuentas(0, 0)
                    oRowD01.Tables(msDebTbG).Rows.Add (oRowD01.Tables(msDebTbG).Rows.Count + 1)
                    oRowD01.Fields("CJRMVD_TOTIMP").ReadOnly = False
                    oRowD01.Fields("CJRMVD_TOTIMP").Value = aSubcuentas(4, nIndice)
                    oRowD01.Fields("CJRMVD_CODDIM").Value = "RLRP"
                    oRowD01.Fields("CJRMVD_DIMDIS").Value = "RLRP"
                    oRowD01.Fields("CJRMVD_CODDIS").Value = aSubcuentas(3, nIndice)
                    oRowD01.Tables(msDebTbG).Rows(oRowD01.Tables(msDebTbG).Rows.Count).Fields("CJRMVG_CUENTA").Value = sCuenta
                    oRowD01.Tables(msDebTbG).Rows(oRowD01.Tables(msDebTbG).Rows.Count).Fields("CJRMVG_CODDIM").Value = "RLRP"
                    oRowD01.Tables(msDebTbG).Rows(oRowD01.Tables(msDebTbG).Rows.Count).Fields("CJRMVG_DIMSUB").Value = "RLRP"
                    oRowD01.Tables(msDebTbG).Rows(oRowD01.Tables(msDebTbG).Rows.Count).Fields("USR_CJRMVG_CODIRL").Value = aSubcuentas(1, nIndice)
                    oRowD01.Tables(msDebTbG).Rows(oRowD01.Tables(msDebTbG).Rows.Count).Fields("USR_CJRMVG_CODIRP").Value = aSubcuentas(2, nIndice)
                    oRowD01.Tables(msDebTbG).Rows(oRowD01.Tables(msDebTbG).Rows.Count).Fields("CJRMVG_IMPORT").ReadOnly = False
                    oRowD01.Tables(msDebTbG).Rows(oRowD01.Tables(msDebTbG).Rows.Count).Fields("CJRMVG_IMPORT").Value = aSubcuentas(4, nIndice)
                    oRowD01.Tables(msDebTbG).Rows(oRowD01.Tables(msDebTbG).Rows.Count).Fields("CJRMVG_SUBCUE").Value = aSubcuentas(3, nIndice)
                Next
            End If
       Next
       
      '--------------------------------------Busco SN--------------------------------------------------
       sCodigoSN = Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_CODISN").Value
       If sCodigoSN = vbNullString Then
            '@ModCom CHAR(2) ,@CodCom VARCHAR(6),@CirCom VARCHAR(6),@CirApl VARCHAR(6),@TipReg VARCHAR(6),@Nrocta VARCHAR(25)
            sSQL2 = "EXEC SP_CG_DLL_BuscoSN '" & sModCom & "','" & sCodCom & "','" & sCirCom & "','" & sCirApl & "','" & sTipReg & "','" & sNroCta & "','" & sTipCpt & "'"
            Set Rs2 = oWizard.DataAccess.OpenResultset(CStr(sSQL2))
            If Not Rs2.EOF Then
                sCodigoSN = ""
                sCodigoSN = Rs2("NuevoCodigoSN").Value   ' Código SN
                sDescSN = Rs2("DescripcionSN").Value
            End If
            If Trim(sCodigoSN) <> "" Then
               Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_CODISN").Value = Trim(sCodigoSN)
               Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_DESCSN").Value = Trim(sDescSN)
            End If
            
        End If
        
    '-----------------------------------AF NUEVO 02-09---------------------------------------------
       sSql = ""
       sSql = "SELECT ISNULL(USR_CJIRVH_PROCES,'N') USR_CJIRVH_PROCES FROM USR_CJIRVH WHERE USR_CJIRVH_NROLOT = " & nNroLot
       Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
       If Not Rs.EOF Then
          If Trim(Rs("USR_CJIRVH_PROCES").Value) = "N" And nNroLot <> 0 Then
             sSQL2 = ""
             sSQL2 = "UPDATE USR_CJIRVH SET USR_CJIRVH_PROCES = 'S' WHERE USR_CJIRVH_NROLOT = " & nNroLot
             Set Rs2 = oWizard.DataAccess.OpenResultset(CStr(sSQL2))
             Rs2.Close
          End If
       End If


    Next

    Set Rs = Nothing

'------z
   
   
   
' ACA RECORRO LA SOLAPA 02  del HAber
    For Each oRowI02 In Who.Table.Rows(1).Tables(msHabTb).Rows

        Who.Table.Rows(1).Tables(msHabTb).Activate
        
        If msDebHab = "" Then
            If msHeadTb = "FCRMVH" Or msHeadTb = "VTRMVH" Then
                AsignoValor Who, oRowI01.Fields("CJRMVI_NROCVT"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
                AsignoValor Who, oRowI01.Fields("CJRMVI_NROSVT"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
            End If
            If msHeadTb = "CORMVH" Or msHeadTb = "PVRMVH" Then
                AsignoValor Who, oRowI01.Fields("CJRMVI_NROCPV"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
                AsignoValor Who, oRowI01.Fields("CJRMVI_NROSPV"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
            End If
        End If
        
        sNroCVT = ""
        sNroCPV = ""
        
        If oRowI02.Fields("CJRMVI_IMPORT").Value = 0 Then
            sSql = " Insert into INTERR (INTERR_NROITM, INTERR_CODTRA, INTERR_FCHMOV, INTERR_FIELDS, INTERR_VALUES, "
            sSql = sSql & " INTERR_ERRMSG , INTERR_FECALT, INTERR_FECMOD, INTERR_HORMOV, INTERR_ULTOPR, INTERR_DEBAJA, "
            sSql = sSql & " INTERR_USERID,INTERR_OALIAS)"
            sSql = sSql & "  Values ((SELECT MAX(IN2.INTERR_NROITM)+1 FROM INTERR AS IN2), 'INTCJM', GETDATE(),"
            sSql = sSql & " 'CJRMVH_CODFOR|USR_CJRMVH_TIPREG|CJRMVH_FCHMOV|CJRMVH_NROCVT|CJRMVH_NROCPV|', "
            sSql = sSql & "'" & Who.Table.Rows(1).Fields(msHeadTb & "_CODFOR").Value & "|"
            sSql = sSql & Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_" & sTipComReg).Value & "|"
            sSql = sSql & CStr(Year(Date)) & Right("00" & CStr(Month(Date)), 2) & Right("00" & CStr(Day(Date)), 2) & "|||',"
            sSql = sSql & "'En el registro " & oRowI02.Fields("CJRMVI_NROITM").Value & " : Imposible grabar con Importe en 0'"
            sSql = sSql & ",  GETDATE(),  GETDATE(),  NULL , 'A', 'N', 'DLL', 'INTERR')"
            Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
            Rs.Close
            Exit Sub
        End If
        nImporte = oRowI02.Fields("CJRMVI_IMPORT").Value
        sModCom = "CJ"
        sCodCom = oRowI02.Fields("CJRMVI_CODFOR").Value
        sCirCom = ""
        sCirApl = ""
        sTipReg = Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_" & sTipComReg).Value
        sTipCpt = oRowI02.Fields("CJRMVI_TIPCPT").Value
        sCodCpt = oRowI02.Fields("CJRMVI_CODCPT").Value
        sCuenta = oRowI02.Fields("CJRMVI_CUENTA").Value
        sTipPro = ""
        sArtCod = ""
        sDeposi = ""
        sSector = ""
        sDepIng = ""
        sSecIng = ""
        sNotros = ""
        nNroLot = oRowI02.Fields("USR_CJRMVI_NROLOT").Value
        nNroIlo = oRowI02.Fields("USR_CJRMVI_NROILO").Value
        sModGen = oRowI02.Fields("USR_CJRMVI_MODGEN").Value
        sComGen = oRowI02.Fields("USR_CJRMVI_COMGEN").Value
        
        sSql = "SELECT ISNULL(USR_CJIRVJ_TIPPRO,'') sTipProIN, ISNULL(USR_CJIRVJ_ARTCOD,'') sArtCodIN, ISNULL(USR_CJIRVJ_CODVIN,'') sNotrosIN,"
        sSql = sSql & " USR_CJIRVJ_NROCVT sNroCVTIN,USR_CJIRVJ_NROCPV sNroCPVIN "
        sSql = sSql & " FROM USR_CJIRVJ "
        sSql = sSql & "  Where USR_CJIRVJ_NROLOT = " & nNroLot
        sSql = sSql & " And USR_CJIRVJ_MODGEN = '" & sModGen & "'"
        sSql = sSql & " And USR_CJIRVJ_COMGEN = '" & sComGen & "'"
        sSql = sSql & " And USR_CJIRVJ_NROILO = " & nNroIlo
        Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
        If Not Rs.EOF Then
            If Trim(Rs("sTipProIN").Value) <> "" Then
               sTipPro = Trim(Rs("sTipProIN").Value)
            End If
            If Trim(Rs("sArtCodIN").Value) <> "" Then
               sArtCod = Trim(Rs("sArtCodIN").Value)
            End If
            If Trim(Rs("sNotrosIN").Value) <> "" Then
               sNotros = Trim(Rs("sNotrosIN").Value)
            End If
           If Trim(Rs("sNroCVTIN").Value) <> "" Then
               sNroCVT = Trim(Rs("sNroCVTIN").Value)
               oRowI02.Fields("CJRMVI_NROCVT").Value = sNroCVT
            End If
            If Trim(Rs("sNroCPVIN").Value) <> "" Then
               sNroCPV = Trim(Rs("sNroCPVIN").Value)
               oRowI02.Fields("CJRMVI_NROCPV").Value = sNroCPV
            End If
        End If
        Rs.Close

        If Trim(sNroCVT) <> "" Then
            sNroCVTCPV = Trim(sNroCVT)
        Else
           If Trim(sNroCPV) <> "" Then
                sNroCVTCPV = Trim(sNroCPV)
           End If
        End If
        sNroCta = sNroCVTCPV 'oRowI01.Fields("CJRMVI_NROCVT").Value
        
        sSql = "EXEC SP_CG_DLL_BuscoCTItem 'DLL', '" & sModCom & "','" & sCodCom & "','" & sCirCom & "','" & sCirApl & "','" & sTipReg & "','" & sNroCta & "','" & sTipCpt & "','" & sCodCpt & "','" & sCuenta & "','" & sTipPro & "','" & sArtCod & "','" & sDeposi & "','" & sSector & "','" & sDepIng & "','" & sSecIng & "', '" & sNotros & "', '" & sRLsVal & "', '" & sRPsVal & "'"
        Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
        If Not Rs.EOF Then
               nCtaMax = Rs("CuentaApertura").Value   ' Cantidad maxima de apertura
               sCuenta = Rs("CuentaValor").Value      ' Cuenta Contable
               sCtaISC = Rs("CuentaImpDim").Value     ' Cuenta es imputable in dimensiones
               sCtaEdi = Rs("CuentaEdita").Value      ' Cuenta es editable
               nSCsMax = Rs("SubCueApertura").Value   ' cantida maxima de apertura de subcuenta por cuenta
               sSCIEdi = "S"                          ' define si edito el importe
               sRLsVal = Rs("RLsValor").Value         ' Valor del RL
               iRLsItm = Rs("RLsItem").Value          ' Item de la USR_VTMRLS del Cliente
               sRPsVal = Rs("RPsValor").Value        ' Valor del RP
               sRLsEdi = Rs("RLsEdita").Value         ' RL es editable
               sRPsEdi = Rs("RPsEdita").Value         ' RP es editable
               sCodVin = ""                           ' Codigo vinulacion siempre es blanco
               sCVsEdi = Rs("CVsEdita").Value         ' CV es editable
               sCVsReq = Rs("CVsRequerido").Value     ' CV es Requerido
               sCVsCnd = Rs("CVsCondicion").Value     ' CV Condicion de validacion
        Else
            Exit Sub
        End If
        Rs.Close
        
'        If Trim(sCuenta) <> "" Then
'            oRowI02.Fields("CJRMVI_CUENTA").Value = sCuenta
'        Else
'            sCuenta = oRowI02.Fields("CJRMVI_CUENTA").Value
'        End If

        
        ' Me fijo si lo debo recuperar de la tabla de interfaces
'        If sRLsVal = "INTERF" Or sRPsVal = "INTERF" Then
'            sSQL = "SELECT USR_CJIRVJ_CODIRL RLsValor, USR_CJIRVJ_CODIRP RPsValor, isnull(USR_CJIRVJ_CODVIN,'') sCodVin "
'            sSQL = sSQL & " FROM USR_CJIRVJ "
'            sSQL = sSQL & "  Where USR_CJIRVJ_NROLOT = " & nNroLot
'            sSQL = sSQL & " And USR_CJIRVJ_MODGEN = '" & sModGen & "'"
'            sSQL = sSQL & " And USR_CJIRVJ_COMGEN = '" & sComGen & "'"
'            sSQL = sSQL & " And USR_CJIRVJ_NROILO = " & nNroIlo
'            Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSQL))
'            If Not Rs.EOF Then
'                If sRLsVal = "INTERF" Then
'                    sRLsVal = Rs("RLsValor").Value
'                End If
'                If sRPsVal = "INTERF" Then
'                    sRPsVal = Rs("RPsValor").Value
'                End If
'                If sRLsVal = "INTERF" And sRPsVal = "INTERF" Then
'                    sCodVin = Rs("sCodVin").Value
'                End If
'            End If
'            Rs.Close
'        End If
        
        sSql = "SELECT ISNULL(USR_CJIRVJ_CODIRL,'') RLsValorIN, ISNULL(USR_CJIRVJ_CODIRP,'') RPsValorIN "
        sSql = sSql & " FROM USR_CJIRVJ "
        sSql = sSql & "  Where USR_CJIRVJ_NROLOT = " & nNroLot
        sSql = sSql & " And USR_CJIRVJ_MODGEN = '" & sModGen & "'"
        sSql = sSql & " And USR_CJIRVJ_COMGEN = '" & sComGen & "'"
        sSql = sSql & " And USR_CJIRVJ_NROILO = " & nNroIlo
        Set Rs = oWizard.DataAccess.OpenResultset(CStr(sSql))
        If Not Rs.EOF Then
            If Trim(Rs("RLsValorIN").Value) <> "" And Trim(sRLsVal) = "" Then
               sRLsVal = Trim(Rs("RLsValorIN").Value)
            End If
            If Trim(Rs("RPsValorIN").Value) <> "" And Trim(sRPsVal) = "" Then
               sRPsVal = Trim(Rs("RPsValorIN").Value)
            End If
        End If
        Rs.Close

        ' aJUSTO LA i
        oRowI02.Fields("CJRMVI_CUENTA").ReadOnly = False
        If Trim(sCuenta) <> "" Then
            If oRowI02.Fields("CJRMVI_CUENTA").Value <> Trim(sCuenta) Then
               oRowI02.Fields("CJRMVI_CUENTA").Value = sCuenta
            End If
        Else
            sCuenta = oRowI02.Fields("CJRMVI_CUENTA").Value
        End If
                
        Erase aSubcuentas
        nIndice = 1
        If Left(sCuenta, 3) <> "CPT" And Left(sRLsVal, 3) <> "CPT" And Left(sRPsVal, 3) <> "CPT" Then
            'Me fijo y actualizo la cantidad de items
            aSubcuentas(0, 0) = nIndice
            
            'actualizo el array
            aSubcuentas(1, nIndice) = sRLsVal
            aSubcuentas(2, nIndice) = sRPsVal
            aSubcuentas(3, nIndice) = sRLsVal & sRPsVal
            aSubcuentas(4, nIndice) = nImporte
            aSubcuentas(5, nIndice) = ""
        
        Else
        
            If Left(sRLsVal, 3) = "CPT" Then
                RecuperoSCSolapa Who, "RL", Right(sRLsVal, 1), sRPsVal
            End If
            
            If Left(sRPsVal, 3) = "CPT" Then
                RecuperoSCSolapa Who, "RP", Right(sRPsVal, 1), sRLsVal
            End If
        
        End If
        
'        With oRowI02.Tables(msHabTbD).Rows
'            oRowI02.Tables(msHabTbD).Rows.Add (oRowI02.Tables(msHabTbD).Rows.Count + 1)
'            Set oRowD02 = oRowI02.Tables(msHabTbD).Rows(oRowI02.Tables(msHabTbD).Rows.Count)

        ' aca voy por la imension
        For Each oRowD02 In oRowI02.Tables(msHabTbD).Rows
            oRowD02.SetError ""
            
            Dim BorroGH As Boolean
            BorroGH = True

            For Each oRowGx In oRowD02.Tables(msHabTbG).Rows
                If oRowGx.Fields("CJRMVG_SUBCUE").Value <> aSubcuentas(3, nIndice) And oRowGx.Fields("CJRMVG_SUBCUE").Value <> "" Then
                    BorroGH = False
                End If
            Next

            
            If (BorroGH) Then

                If oRowD02.Tables(msHabTbG).Rows.Count > 0 Then
                    oRowD02.Tables(msHabTbG).Rows.Clear
                End If
                For nIndice2 = 1 To aSubcuentas(0, 0)
                    oRowD02.Tables(msHabTbG).Rows.Add (oRowD02.Tables(msHabTbG).Rows.Count + 1)
                    oRowD02.Fields("CJRMVD_TOTIMP").ReadOnly = False
                    oRowD02.Fields("CJRMVD_TOTIMP").Value = aSubcuentas(4, nIndice)
                    oRowD02.Fields("CJRMVD_CODDIM").Value = "RLRP"
                    oRowD02.Fields("CJRMVD_DIMDIS").Value = "RLRP"
                    oRowD02.Fields("CJRMVD_CODDIS").Value = aSubcuentas(3, nIndice)
                    oRowD02.Tables(msHabTbG).Rows(oRowD02.Tables(msHabTbG).Rows.Count).Fields("CJRMVG_CUENTA").Value = sCuenta
                    oRowD02.Tables(msHabTbG).Rows(oRowD02.Tables(msHabTbG).Rows.Count).Fields("CJRMVG_CODDIM").Value = "RLRP"
                    oRowD02.Tables(msHabTbG).Rows(oRowD02.Tables(msHabTbG).Rows.Count).Fields("CJRMVG_DIMSUB").Value = "RLRP"
                    oRowD02.Tables(msHabTbG).Rows(oRowD02.Tables(msHabTbG).Rows.Count).Fields("USR_CJRMVG_CODIRL").Value = aSubcuentas(1, nIndice)
                    oRowD02.Tables(msHabTbG).Rows(oRowD02.Tables(msHabTbG).Rows.Count).Fields("USR_CJRMVG_CODIRP").Value = aSubcuentas(2, nIndice)
                    oRowD02.Tables(msHabTbG).Rows(oRowD02.Tables(msHabTbG).Rows.Count).Fields("CJRMVG_IMPORT").ReadOnly = False
                    oRowD02.Tables(msHabTbG).Rows(oRowD02.Tables(msHabTbG).Rows.Count).Fields("CJRMVG_IMPORT").Value = aSubcuentas(4, nIndice)
                    oRowD02.Tables(msHabTbG).Rows(oRowD02.Tables(msHabTbG).Rows.Count).Fields("CJRMVG_SUBCUE").Value = aSubcuentas(3, nIndice)
                    
                Next
            End If
        Next
        'End With
       '--------------------------------------Busco SN--------------------------------------------------
       sCodigoSN = Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_CODISN").Value
       If sCodigoSN = vbNullString Then
            '@ModCom CHAR(2) ,@CodCom VARCHAR(6),@CirCom VARCHAR(6),@CirApl VARCHAR(6),@TipReg VARCHAR(6),@Nrocta VARCHAR(25)
            sSQL2 = "EXEC SP_CG_DLL_BuscoSN '" & sModCom & "','" & sCodCom & "','" & sCirCom & "','" & sCirApl & "','" & sTipReg & "','" & sNroCta & "','" & sTipCpt & "'"
            Set Rs2 = oWizard.DataAccess.OpenResultset(CStr(sSQL2))
            If Not Rs2.EOF Then
                sCodigoSN = ""
                sCodigoSN = Rs2("NuevoCodigoSN").Value   ' Código SN
                sDescSN = Rs2("DescripcionSN").Value
            End If
            If Trim(sCodigoSN) <> "" Then
               Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_CODISN").Value = Trim(sCodigoSN)
               Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_DESCSN").Value = Trim(sDescSN)
            End If
            
       End If
       
    Next
    
    Set Rs = Nothing

End Sub
'*******************************************
'                            Recupero subcuenta en base a los conceptos de la solapa indicada
Sub RecuperoSCSolapa(Who As Object, sConceptoRLRP As String, sSolapaX As String, sConceptoFijo As String)
'*******************************************
    Dim sQueSolapa As String
    Dim sQueSolapaD As String
    Dim sQueSolapaG As String

    Dim oRowI As Object
    Dim oRowG As Object
    Dim nIndice As Integer
    Dim nUltimoOk As Integer
    
    nIndice = 0
    
    ' me fijo que solapa hacer y la seteo
    If sSolapaX = "D" Then
        sQueSolapa = msDebTb
        sQueSolapaD = msDebTbD
        sQueSolapaG = msDebTbG

        End If
    If sSolapaX = "H" Then
        sQueSolapa = msHabTb
        sQueSolapaD = msHabTbD
        sQueSolapaG = msHabTbG

    End If
    
    For Each oRowI In Who.Table.Rows(1).Tables(sQueSolapa).Rows
        For Each oRowG In oRowI.Tables(sQueSolapaD).Rows(1).Tables(sQueSolapaG).Rows
            
            'Seteo en cual estoy parado y voy agregando
            nIndice = nIndice + 1
            aSubcuentas(0, 0) = nIndice
            
            'Me fijo
            If sConceptoRLRP = "RL" Then
                 aSubcuentas(1, nIndice) = oRowG.Fields("USR_CJRMVG_CODIRL").Value
                If Left(sConceptoFijo, 3) <> "CPT" Then
                    aSubcuentas(2, nIndice) = sConceptoFijo
                End If
            Else
                aSubcuentas(2, nIndice) = oRowG.Fields("USR_CJRMVG_CODIRP").Value
                If Left(sConceptoFijo, 3) <> "CPT" Then
                    aSubcuentas(1, nIndice) = sConceptoFijo
                End If
            End If
           
            aSubcuentas(4, nIndice) = aSubcuentas(1, nIndice) & aSubcuentas(2, nIndice)
    
        Next
    Next

End Sub

Private Sub DefinoTablas(Who As Object)
    msMainTb = Who.Table.Name
    msHeadTb = msMainTb
    sTipComReg = "TIPCOM"
    Select Case msMainTb
        Case "CJRMVH"
            msDebTb = "CJRMVI01"
            msHabTb = "CJRMVI02"
            msDebTbD = "CJRMVD01"
            msHabTbD = "CJRMVD02"
            msDebTbG = "CJRMVG01"
            msHabTbG = "CJRMVG02"
            sTipComReg = "TIPREG"
        Case "VTRRCH"
            msHeadTb = "VTRMVH"
            msDebTb = "VTRRCC04"
            msHabTb = "VTRRCC05"
            msDebTbD = "CJRMVD03"
            msHabTbD = "CJRMVD04"
            msDebTbG = "CJRMVG03"
            msHabTbG = "CJRMVG04"
            
        Case "PVRPPH"
            msHeadTb = "PVRMVH"
            msDebTb = "PVRRCC04"
            msHabTb = "PVRRCC05"
            msDebTbD = "CJRMVD05"
            msHabTbD = "CJRMVD06"
            msDebTbG = "CJRMVG05"
            msHabTbG = "CJRMVG06"
        Case "FCRMVH"
            msDebTb = "FCRMVI08"
            msHabTb = "FCRMVI09"
            msDebTbD = "CJRMVD07"
            msHabTbD = "CJRMVD08"
            msDebTbG = "CJRMVG07"
            msHabTbG = "CJRMVG08"
        Case "CORMVH"
            msDebTb = "CORMVI08"
            msHabTb = "CORMVI09"
            msDebTbD = "CJRMVD09"
            msHabTbD = "CJRMVD10"
            msDebTbG = "CJRMVG09"
            msHabTbG = "CJRMVG10"
        Case "FCRMVHUPD"
            msHeadTb = "FCRMVH"
            msDebTb = "FCRMVI8UPD"
            msHabTb = "FCRMVI9UPD"
            msDebTbD = "CJRMVD11"
            msHabTbD = "CJRMVD12"
            msDebTbG = "CJRMVG11"
            msHabTbG = "CJRMVG12"
        Case "CORMVHUPD"
            msHeadTb = "CORMVH"
            msDebTb = "CORMVI8UPD"
            msHabTb = "CORMVI9UPD"
            msDebTbD = "CJRMVD13"
            msHabTbD = "CJRMVD14"
            msDebTbG = "CJRMVG13"
            msHabTbG = "CJRMVG14"
        
    End Select
    mPasaValImport = False
End Sub


Private Sub DefinoModulo(Who As Object)
    msModulo = CStr(Who.ModuleName) 'Defino el Módulo de la aplicación (CJ, etc)
    If Not bControladoCircuito Then
        DefineSiEsAnulacionTotal Who
    End If

End Sub


Private Function ValidoTablas(Who As Object) As Integer
    Select Case Who.CurrentTable
        Case msDebTb
            HabilitaNROCVT Who, msDebTb
    End Select

End Function


Private Function ValidoCampos(Who As Object) As Integer
           
Dim test As String, Apertura As Integer
test = Who.CurrentField
    
    Select Case Who.CurrentField
        Case "USR_" & msHeadTb & "_" & sTipComReg
            RecargaConceptos Who
            If msMainTb <> "CJRMVH" Then
                Apertura = BuscaAperturaCtble(Who, msHabTb, msHabTbD, msHabTbG)
                Apertura = BuscaAperturaCtble(Who, msDebTb, msDebTbD, msDebTbG)
            End If
        Case "CJRMVI_CODCPT"
            If Who.CurrentTable = msDebTb Then
                HabilitaNROCVT Who, msDebTb
            End If
        Case "CJRMVI_IMPORT", "CJRMVI_NROINT", "CJRMVI_IMPUSS"
                Valida_Import Who
            'End If
        Case "CJRMVI_NROCVT"
             'If who.CurrentTable = msDebTb Then
             '   CompletaNROCVT who, msDebTb
             'End If
             'If who.CurrentTable = msHabTb Then
             '   CompletaNROCVT who, msHabTb
             'End If
        Case "CJRMVI_NROCPV"
             'If who.CurrentTable = msDebTb Then
             '   CompletaNROCPV who, msDebTb
             'End If
             'If who.CurrentTable = msHabTb Then
             '   CompletaNROCPV who, msHabTb
             'End If
         Case "CJRMVG_IMPORT"
             ValidaImportDConI Who
         Case "USR_CJRMVG_CODIRL", "USR_CJRMVG_CODIRP"
                ReasignaSC Who
             
    End Select

End Function


Private Function HabilitaNROCVT(Who As Object, sAlias As String)
    Dim oRowI As Object
    
      With Who.Table.Rows(1).Tables(sAlias)
            If .Rows.Count > 0 Then
            For Each oRowI In .Rows
                 With oRowI
                    oRowI.Fields("CJRMVI_NROCVT").ReadOnly = False
                 End With
            Next
            End If
      End With

End Function


Private Function RecargaConceptos(Who As Object)
    Dim iValidaCarga As Integer
    Dim sSql As String
    Dim Recset As rdoResultset
    
    
    msDebHab = ""
    If Trim(Who.Table.Rows(1).Fields(msHeadTb & "_CODFOR").Value = "") Then
       Exit Function
    End If
    If msHeadTb = "CJRMVH" Then
        Who.Table.Rows(1).Fields("CJRMVH_NROCVT").ReadOnly = False
        Who.Table.Rows(1).Fields("CJRMVH_NROCPV").ReadOnly = False
        Who.Table.Rows(1).Fields("CJRMVH_NOMBVT").ReadOnly = False
        Who.Table.Rows(1).Fields("CJRMVH_NOMBPV").ReadOnly = False
        Who.Table.Rows(1).Fields("CJRMVH_NOMBVT").Value = ""
        Who.Table.Rows(1).Fields("CJRMVH_NOMBPV").Value = ""
    End If
    
    sSql = vbNullString
    sSql = "SELECT ISNULL(USR_GRMTRH_NROCVT,'') USR_GRMTRH_NROCVT,ISNULL(USR_GRMTRH_NROCPV,'') USR_GRMTRH_NROCPV, "
    sSql = sSql & " ISNULL(USR_GRMTRH_DEBHAB,'') USR_GRMTRH_DEBHAB "
    sSql = sSql & " From USR_GRMTRH "
    sSql = sSql & " WHERE USR_GRMTRH_MODREG ='CJ' AND "
    sSql = sSql & " USR_GRMTRH_TIPREG='" & Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_" & sTipComReg).Value & "' "
    Set Recset = Who.DataAccess.OpenResultset(sSql, 1)
    If Not Recset.EOF Then
        If Trim(Recset("USR_GRMTRH_NROCVT")) <> "" Then
            AsignoValor Who, Who.Table.Rows(1).Fields("CJRMVH_NROCVT"), Trim(Recset("USR_GRMTRH_NROCVT")), Who.Table.Rows(1).Fields("CJRMVH_NROCVT").ReadOnly
            
            If msHeadTb = "CJRMVH" Then
                Who.Table.Rows(1).Fields("CJRMVH_NROCVT").ReadOnly = True
                Who.Table.Rows(1).Fields("CJRMVH_NOMBVT").ReadOnly = True
            End If
        End If
        If Trim(Recset("USR_GRMTRH_NROCPV")) <> "" Then
            AsignoValor Who, Who.Table.Rows(1).Fields("CJRMVH_NROCPV"), Trim(Recset("USR_GRMTRH_NROCPV")), Who.Table.Rows(1).Fields("CJRMVH_NROCPV").ReadOnly
          
            If msHeadTb = "CJRMVH" Then
                Who.Table.Rows(1).Fields("CJRMVH_NROCPV").ReadOnly = True
                Who.Table.Rows(1).Fields("CJRMVH_NOMBPV").ReadOnly = True
            End If
        End If
        If Trim(Recset("USR_GRMTRH_DEBHAB")) <> "" Then
            msDebHab = Trim(Recset("USR_GRMTRH_DEBHAB"))
        End If
    End If
    Recset.Close
    Set Recset = Nothing
'-------------Recargar Conceptos de acuerdo a lo definido en el cpte de Tesoreria
'-------------Primero borro los items----------------------------------

 iValidaCarga = 0
 iValidaCarga = CargoItems(Who, msDebTb, msDebTbD, msDebTbG)
 If iValidaCarga = 0 Then
    iValidaCarga = CargoItems(Who, msHabTb, msHabTbD, msHabTbG)
 End If
 
 If iValidaCarga <> 0 Then
    Who.SetError "GR000068", "Error al Cargar Items de Tipo de Registracion"
 End If
  
 
End Function
Private Function BorroItems(Who As Object, sAlias As String)
Dim oRowI As Object
Dim oField As Object
     
    With Who.Table.Rows(1).Tables(sAlias)
          .Rows.Clear
    End With
End Function



Private Function CargoItems(Who As Object, sAlias As String, sAliasD As String, sAliasG As String) As Integer
    Dim sSqlComp As String
    Dim sSql As String
    Dim sSQL2 As String
    
    Dim Recset As rdoResultset
    Dim Recset2 As rdoResultset
    Dim RecsetComp As rdoResultset
    
    Dim Filact As Long
    Dim oRowI As Object
    Dim HabAgr As Boolean
    Dim dImporte As Double
    Dim sDebHab As String

    
    sSqlComp = "SELECT isnull("
    If sAlias = msDebTb Then
        sSqlComp = sSqlComp & " USR_GRCCBH_CARITD"
    Else
        sSqlComp = sSqlComp & " USR_GRCCBH_CARITH"
    End If
    sSqlComp = sSqlComp & ",'N') CARGAITM FROM GRCCBH WHERE GRCCBH_MODCOM = 'CJ' AND GRCCBH_CODCOM = '"
    sSqlComp = sSqlComp + Who.Table.Rows(1).Fields(msHeadTb & "_CODFOR").Value + "'"
    Set RecsetComp = Who.DataAccess.OpenResultset(sSqlComp, 1)
    If Not RecsetComp.EOF Then
        If RecsetComp("CARGAITM").Value = "N" Then
            Set RecsetComp = Nothing
            CargoItems = 0
            Exit Function
        End If
    End If
    
    Set RecsetComp = Nothing
    
    BorroItems Who, sAlias
    Who.IntoEvent = False
    Filact = 0
    HabAgr = False    'ojo hay un concepto que lo ponga en true, ya no deja agregar items en la solapa

With Who.Table.Rows(1).Tables(sAlias)
         '.Activate
         Filact = .Rows.Count
          .Rows.MaxRows = 50
         sSql = vbNullString
         sSql = "SELECT USR_GRCCRG_TIPCPT,USR_GRCCRG_CODCPT, isnull(USR_GRCCRG_HABAGR,'N') USR_GRCCRG_HABAGR, "
         sSql = sSql & " ISNULL(USR_GRCCRG_TIPRET,'') USR_GRCCRG_TIPRET,ISNULL(USR_GRCCRG_CODRET,'') USR_GRCCRG_CODRET  "
         sSql = sSql & " From USR_GRCCRG "
         sSql = sSql & " WHERE USR_GRCCRG_MODCOM='CJ' AND "
         sSql = sSql & " USR_GRCCRG_CODCOM='" & Who.Table.Rows(1).Fields(msHeadTb & "_CODFOR").Value & "' AND "
         sSql = sSql & " USR_GRCCRG_MODREG='CJ' AND "
         sSql = sSql & " USR_GRCCRG_TIPREG='" & Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_" & sTipComReg).Value & "' "
         If sAlias = msDebTb Then
            sSql = sSql & " AND USR_GRCCRG_DEBHAB='D'"
         Else
            sSql = sSql & " AND USR_GRCCRG_DEBHAB='H'"
         End If
         Set Recset = Who.DataAccess.OpenResultset(sSql, 1)
         Do While Not Recset.EOF
            Filact = Filact + 1
            .Rows.Add (Filact)
            Set oRowI = .Rows(Filact)
            With oRowI
                AsignoValor Who, .Fields("CJRMVI_TIPCPT"), Recset("USR_GRCCRG_TIPCPT").Value, .Fields("CJRMVI_TIPCPT").ReadOnly
                AsignoValor Who, .Fields("CJRMVI_CODCPT"), Trim(Recset("USR_GRCCRG_CODCPT").Value), .Fields("CJRMVI_CODCPT").ReadOnly
                                
                If Trim(Recset("USR_GRCCRG_TIPRET").Value) <> "" Then
                    AsignoValor Who, .Fields("USR_CJRMVI_TIPRET"), Trim(Recset("USR_GRCCRG_TIPRET").Value), .Fields("USR_CJRMVI_TIPRET").ReadOnly
                End If
                If Trim(Recset("USR_GRCCRG_CODRET").Value) <> "" Then
                    AsignoValor Who, .Fields("USR_CJRMVI_CODRET"), Trim(Recset("USR_GRCCRG_CODRET").Value), .Fields("USR_CJRMVI_CODRET").ReadOnly
                End If
                '-------------------Bajo Nro Cliente / Nro Proveedor de Tipo Registración----------------------------------
                '-------OJO ES CRUZADO---SI DEFINE DEBE BUSCA EL NRO CLIENTE/PROV DEL HABER EN LA CTA DE TESORERIA---------
                oRowI.Fields("CJRMVI_NROCVT").ReadOnly = False
                oRowI.Fields("CJRMVI_NROSVT").ReadOnly = False
                oRowI.Fields("CJRMVI_NOMBVT").ReadOnly = False
                oRowI.Fields("CJRMVI_NROCPV").ReadOnly = False
                oRowI.Fields("CJRMVI_NROSPV").ReadOnly = False
                oRowI.Fields("CJRMVI_NOMBPV").ReadOnly = False
                'oRowI.Fields("CJRMVI_DOCFIR").ReadOnly = False
                '--------------------------------INICIALIZO PORQUE BAJA DE LA CABECERA-------------------------------------
                oRowI.Fields("CJRMVI_NROCVT").Value = vbNullString
                oRowI.Fields("CJRMVI_NROSVT").Value = vbNullString
                oRowI.Fields("CJRMVI_NOMBVT").Value = " "
                'oRowI.Fields("CJRMVI_DOCFIR").Value = vbNullString
                oRowI.Fields("CJRMVI_NROCPV").Value = vbNullString
                oRowI.Fields("CJRMVI_NROSPV").Value = vbNullString
                oRowI.Fields("CJRMVI_NOMBPV").Value = " "
                oRowI.Fields("CJRMVI_FCHAUX").SetError ""
                oRowI.SetError ""

                If oRowI.Fields("CJRMVI_IMPORT").ReadOnly = False Then
                    AsignoValor Who, oRowI.Fields("CJRMVI_IMPORT"), CStr(RecuperaImporteNetoPagoCob(Who, sAlias, Filact)), False
                End If
                If msDebHab = "" Then
                    If msHeadTb = "FCRMVH" Or msHeadTb = "VTRMVH" Then
                        oRowI.Fields("CJRMVI_NROCVT").Value = Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value
                        oRowI.Fields("CJRMVI_NROSVT").Value = Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value
                    End If
                    If msHeadTb = "CORMVH" Or msHeadTb = "PVRMVH" Then
                        oRowI.Fields("CJRMVI_NROCPV").Value = Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value
                        oRowI.Fields("CJRMVI_NROSPV").Value = Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value
                    End If
                End If
                If msDebHab = "" Then
                    If msHeadTb = "FCRMVH" Or msHeadTb = "VTRMVH" Then
                        oRowI.Fields("CJRMVI_NROCVT").Value = Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value
                       oRowI.Fields("CJRMVI_NROSVT").Value = Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value
                    End If
                    If msHeadTb = "CORMVH" Or msHeadTb = "PVRMVH" Then
                        oRowI.Fields("CJRMVI_NROCPV").Value = Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value
                        oRowI.Fields("CJRMVI_NROSPV").Value = Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value
                    End If
                End If
                
                If msDebHab = "D" Then
                    If sAlias = msDebTb Then
                        If Trim(Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value) <> "" Then
                            AsignoValor Who, oRowI.Fields("CJRMVI_NROCVT"), Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value, oRowI.Fields("CJRMVI_NROCVT").ReadOnly
                        End If
                        If Trim(Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value) <> "" Then
                            AsignoValor Who, oRowI.Fields("CJRMVI_NROCPV"), Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value, oRowI.Fields("CJRMVI_NROCPV").ReadOnly
                        End If
                    Else
                        sSQL2 = vbNullString
                        sSQL2 = "SELECT ISNULL(USR_CJTCTA_NROCVT,'') USR_CJTCTA_NROCVT,ISNULL(USR_CJTCTA_NROCPV,'') USR_CJTCTA_NROCPV "
                        sSQL2 = sSQL2 & " From CJTCTA INNER JOIN GRCCOH ON GRCCOH_CJCTACTE = CJTCTA_CTACTE "
                        sSQL2 = sSQL2 & " WHERE GRCCOH_TIPCPT = '" & oRowI.Fields("CJRMVI_TIPCPT").Value
                        sSQL2 = sSQL2 & "' AND GRCCOH_CODCPT='" & oRowI.Fields("CJRMVI_CODCPT").Value & "'"
                        Set Recset2 = Who.DataAccess.OpenResultset(sSQL2, 1)
                        If Not Recset2.EOF Then
                            If Trim(Recset2("USR_CJTCTA_NROCVT").Value) <> "" Then
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROCVT"), Recset2("USR_CJTCTA_NROCVT").Value, oRowI.Fields("CJRMVI_NROCVT").ReadOnly
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROSVT"), Recset2("USR_CJTCTA_NROCVT").Value, oRowI.Fields("CJRMVI_NROSVT").ReadOnly
                            End If
                            If Trim(Recset2("USR_CJTCTA_NROCPV").Value) <> "" Then
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROCPV"), Recset2("USR_CJTCTA_NROCPV").Value, oRowI.Fields("CJRMVI_NROCPV").ReadOnly
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROSPV"), Recset2("USR_CJTCTA_NROCPV").Value, oRowI.Fields("CJRMVI_NROSPV").ReadOnly
                            End If
                        End If
                        If Trim(oRowI.Fields("CJRMVI_NROCVT").Value) = "" And _
                           Trim(oRowI.Fields("CJRMVI_NROCPV").Value) = "" Then
                            If Trim(Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value) <> "" Then
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROCVT"), Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value, oRowI.Fields("CJRMVI_NROCVT").ReadOnly
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROSVT"), Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value, oRowI.Fields("CJRMVI_NROSVT").ReadOnly
                            End If
                            If Trim(Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value) <> "" Then
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROCPV"), Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value, oRowI.Fields("CJRMVI_NROCPV").ReadOnly
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROSPV"), Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value, oRowI.Fields("CJRMVI_NROSPV").ReadOnly
                            End If
                        End If
                    End If
                End If

                If msDebHab = "H" Then
                    If sAlias = msHabTb Then
                        If Trim(Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value) <> "" Then
                            AsignoValor Who, oRowI.Fields("CJRMVI_NROCVT"), Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value, oRowI.Fields("CJRMVI_NROCVT").ReadOnly
                            AsignoValor Who, oRowI.Fields("CJRMVI_NROSVT"), Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value, oRowI.Fields("CJRMVI_NROSVT").ReadOnly
                        End If
                        If Trim(Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value) <> "" Then
                            AsignoValor Who, oRowI.Fields("CJRMVI_NROCPV"), Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value, oRowI.Fields("CJRMVI_NROCPV").ReadOnly
                            AsignoValor Who, oRowI.Fields("CJRMVI_NROSPV"), Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value, oRowI.Fields("CJRMVI_NROSPV").ReadOnly
                        End If
                    Else
                        sSQL2 = vbNullString
                        sSQL2 = "SELECT ISNULL(USR_CJTCTA_NROCVT,'') USR_CJTCTA_NROCVT,ISNULL(USR_CJTCTA_NROCPV,'') USR_CJTCTA_NROCPV "
                        sSQL2 = sSQL2 & " From CJTCTA INNER JOIN GRCCOH ON GRCCOH_CJCTACTE = CJTCTA_CTACTE "
                        sSQL2 = sSQL2 & " WHERE GRCCOH_TIPCPT = '" & oRowI.Fields("CJRMVI_TIPCPT").Value
                        sSQL2 = sSQL2 & "' AND GRCCOH_CODCPT='" & oRowI.Fields("CJRMVI_CODCPT").Value & "'"
                        Set Recset2 = Who.DataAccess.OpenResultset(sSQL2, 1)
                        If Not Recset2.EOF Then
                            If Trim(Recset2("USR_CJTCTA_NROCVT").Value) <> "" Then
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROCVT"), Recset2("USR_CJTCTA_NROCVT").Value, oRowI.Fields("CJRMVI_NROCVT").ReadOnly
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROSVT"), Recset2("USR_CJTCTA_NROCVT").Value, oRowI.Fields("CJRMVI_NROSVT").ReadOnly
                            End If
                            If Trim(Recset2("USR_CJTCTA_NROCPV").Value) <> "" Then
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROCPV"), Recset2("USR_CJTCTA_NROCPV").Value, oRowI.Fields("CJRMVI_NROSPV").ReadOnly
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROSPV"), Recset2("USR_CJTCTA_NROCPV").Value, oRowI.Fields("CJRMVI_NROSPV").ReadOnly
                            End If
                        End If
                        If Trim(oRowI.Fields("CJRMVI_NROCVT").Value) = "" And _
                           Trim(oRowI.Fields("CJRMVI_NROCPV").Value) = "" Then
                            If Trim(Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value) <> "" Then
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROCVT"), Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value, oRowI.Fields("CJRMVI_NROCVT").ReadOnly
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROSVT"), Who.Table.Rows(1).Fields("CJRMVH_NROCVT").Value, oRowI.Fields("CJRMVI_NROSVT").ReadOnly
                            End If
                            If Trim(Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value) <> "" Then
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROCPV"), Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value, oRowI.Fields("CJRMVI_NROCPV").ReadOnly
                                AsignoValor Who, oRowI.Fields("CJRMVI_NROSPV"), Who.Table.Rows(1).Fields("CJRMVH_NROCPV").Value, oRowI.Fields("CJRMVI_NROSPV").ReadOnly
                            End If
                        End If
                    End If
                End If

                oRowI.Fields("CJRMVI_NROCVT").ReadOnly = True
                oRowI.Fields("CJRMVI_NROSVT").ReadOnly = True
                oRowI.Fields("CJRMVI_NOMBVT").ReadOnly = True
                oRowI.Fields("CJRMVI_NROCPV").ReadOnly = True
                oRowI.Fields("CJRMVI_NROSPV").ReadOnly = True
                oRowI.Fields("CJRMVI_NOMBPV").ReadOnly = True
                oRowI.Fields("CJRMVI_FCHAUX").SetError ""
                oRowI.SetError ""
                'oRowI.Fields("CJRMVI_DOCFIR").ReadOnly = True
                '----------------------------------------------------------------

            End With
            If HabAgr = False And Recset("USR_GRCCRG_HABAGR").Value = "S" Then
                HabAgr = True
            End If
                
            Recset.MoveNext
            
         Loop
         If HabAgr = False Then
            .Rows.MaxRows = Filact
         End If
    End With
 
    Recset.Close
    Set Recset = Nothing
    Set oRowI = Nothing



    CargoItems = 0
'who.IntoEvent = True
End Function

Private Function Valida_Import(Who As Object) As Integer
    Dim Valida As Integer
    Dim oRow As Object
    Dim sCurTabDH As String
    Dim sCurTabDHD As String
    Dim sCurTabDHG As String
    sCurTabDH = vbNullString
    Select Case Who.CurrentTable
        Case msDebTb
            sCurTabDH = msDebTb
            sCurTabDHD = msDebTbD
            sCurTabDHG = msDebTbG
        Case msHabTb
            sCurTabDH = msHabTb
            sCurTabDHD = msHabTbD
            sCurTabDHG = msHabTbG
    End Select
    
    If Trim(sCurTabDH) <> "" Then
        With Who.Table.Rows(1).Tables(sCurTabDH)
               Set oRow = .Rows(.CurrentRow)
               With oRow
                    If .Fields("CJRMVI_IMPORT").Value <> 0 Then
                        Who.IntoEvent = True
                        mPasaValImport = True
                        Valida = BuscaAperturaCtble(Who, sCurTabDH, sCurTabDHD, sCurTabDHG)
                                                'Who.IntoEvent = True
                        mPasaValImport = False
                    End If
               End With
        End With
   End If
End Function


'*******************************************
'              GESTION DE  cambio de RL o RP
Private Sub ReasignaSC(Who As Object)
'*******************************************
Dim sCurTabDH, sCurTabDHD, sCurTabDHG
    Select Case Who.CurrentTable
        Case msDebTbG
            sCurTabDH = msDebTb
            sCurTabDHD = msDebTbD
            sCurTabDHG = msDebTbG
        Case msHabTbG
            sCurTabDH = msHabTb
            sCurTabDHD = msHabTbD
            sCurTabDHG = msHabTbG
    End Select


    
    'si cambia manualmente IDRL o IDRP  armo la subcuenta
    With Who.Table.Rows(1).Tables(sCurTabDH)
        With .Rows(.CurrentRow).Tables(sCurTabDHD)
            With .Rows(.CurrentRow).Tables(sCurTabDHG)
                If .Rows(.CurrentRow).Fields("CJRMVG_SUBCUE").Value <> Right(.Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRL").Value, 5) + Right(.Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRP").Value, 5) Then
                    Who.IntoEvent = False
                    
                    If Right(.Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRL").Value, 5) <> "" And Right(.Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRP").Value, 5) <> "" Then
                        AsignoValor Who, .Rows(.CurrentRow).Fields("CJRMVG_SUBCUE"), Right(.Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRL").Value, 5) + Right(.Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRP").Value, 5), False
                    Else
                        AsignoValor Who, .Rows(.CurrentRow).Fields("CJRMVG_SUBCUE"), "", False
                    End If
                    
                    If Not bEsAnuTot Then
                     
                        .Rows(.CurrentRow).Fields("CJRMVG_SUBCUE").ReadOnly = True
                    Else
                        
                        Dim sModapl, sCodapl, sNroapl, sGNroItm, sGNroItd, sGDebHab, sGCodDim, sGNroIts, sGSubCue, sSql, Rs

                        
                        sModapl = Who.Table.Rows(1).Tables(1).Rows(1).Fields(Who.Table.Rows(1).Tables(1).Name & "_MODAPL").SQLValue
                        sCodapl = Who.Table.Rows(1).Tables(1).Rows(1).Fields(Who.Table.Rows(1).Tables(1).Name & "_CODAPL").SQLValue
                        sNroapl = Who.Table.Rows(1).Tables(1).Rows(1).Fields(Who.Table.Rows(1).Tables(1).Name & "_NROAPL").SQLValue
                        sGNroItm = .Rows(.CurrentRow).Fields("CJRMVG_NROITM").SQLValue
                        sGNroItd = .Rows(.CurrentRow).Fields("CJRMVG_NROITD").SQLValue
                        sGNroIts = .Rows(.CurrentRow).Fields("CJRMVG_NROITS").SQLValue
                        sGDebHab = .Rows(.CurrentRow).Fields("CJRMVG_DEBHAB").SQLValue
                        sGCodDim = .Rows(.CurrentRow).Fields("CJRMVG_CODDIM").SQLValue
                        
                        sSql = "SELECT CJRMVG_SUBCUE FROM CJRMVG WHERE CJRMVG_MODFOR = 'CJ'"
                        sSql = sSql & " AND CJRMVG_CODFOR = " & sCodapl
                        sSql = sSql & " AND CJRMVG_NROFOR = " & sNroapl
                        sSql = sSql & " AND CJRMVG_NROITM = " & sGNroItm
                        sSql = sSql & " AND CJRMVG_NROITD = " & sGNroItd
                        sSql = sSql & " AND CJRMVG_NROITS = " & sGNroIts
                        sSql = sSql & " AND CJRMVG_DEBHAB = " & sGDebHab
                        sSql = sSql & " AND CJRMVG_CODDIM = " & sGCodDim
                        
                        Set Rs = Who.DataAccess.OpenResultset(CStr(sSql))
                            
                        If Not Rs.EOF Then
                            sGSubCue = Rs("CJRMVG_SUBCUE").Value
                        End If



                        sSql = "SELECT * FROM DBO.FN_CG_SEL_SUBCUE_IDRL_IDRP (" & sGCodDim & ",'" & sGSubCue & "')"
                        Set Rs = Who.DataAccess.OpenResultset(CStr(sSql))
                            
                        If Not Rs.EOF Then
                        Who.Validating = False
                            .Rows(.CurrentRow).Fields("CJRMVG_SUBCUE").Value = Rs("SUBCUE").Value
                            .Rows(.CurrentRow).Fields("CJRMVG_SUBCUE").DescriptionValue = Rs("DS_SUBCUE").Value
                            .Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRL").Value = Rs("IDRL").Value
                            .Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRL").DescriptionValue = Rs("DS_IDRL").Value
                            .Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRP").Value = Rs("IDRP").Value
                            .Rows(.CurrentRow).Fields("USR_CJRMVG_CODIRP").DescriptionValue = Rs("DS_IDRP").Value
'                            MsgBox ("A")
                       End If


                    End If
                    
                End If
            End With
        End With
    End With

End Sub


Private Function BuscaAperturaCtble(Who As Object, sAlias As String, sAliasD As String, sAliasG As String) As Integer
    If bEsAnuTot = False Or bReady = True Then
        
        Dim oRowI As Object
        Dim oRowD As Object
        Dim sSql As String
        Dim sSQL2 As String
        Dim Rs As rdoResultset
        Dim Rs2 As rdoResultset
        Dim nIndice As Integer
        Dim nIndice2 As Integer
        Dim nNroLot As Integer
        Dim nNroIlo As Integer
        Dim sModCom As String
        Dim sCodCom As String
        Dim sCirCom As String
        Dim sCirApl As String
        Dim sTipReg As String
        Dim sNroCta As String
        Dim sTipCpt As String
        Dim sCodCpt As String
        Dim sCuenta As String
        Dim sTipPro As String
        Dim sArtCod As String
        Dim sDeposi As String
        Dim sSector As String
        Dim sDepIng As String
        Dim sSecIng As String
        Dim sNotros As String
        Dim sCodigoSN As String
        Dim sDescSN As String
        Dim nItmSub As Integer
        Dim nImporte As Double
        Dim nCtaMax As Double
        Dim sCtaISC As String
        Dim sCtaEdi As String
        Dim nSCsMax As Double
        Dim sSCIEdi As String
        Dim sRLsVal As String
        Dim iRLsItm As Integer
        Dim sRPsVal As String
        Dim sRLsEdi As String
        Dim sRPsEdi As String
        Dim sCodVin As String
        Dim sCVsEdi As String
        Dim sCVsReq As String
        Dim sCVsCnd As String
        
        For Each oRowI In Who.Table.Rows(1).Tables(sAlias).Rows
            'Who.IntoEvent = False
            If msMainTb = "CJRMVH" Then 'condición agregada por BFM 14-05-2019
 'comentado 20190903               Who.Validating = True
            End If
            
            If oRowI.Fields("CJRMVI_IMPORT").Value <> 0 Then
            
                If msDebHab = "" Then
                    If msHeadTb = "FCRMVH" Or msHeadTb = "VTRMVH" Then
                        If Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value <> "" Then
                            If msMainTb = "CJRMVH" Then
 'comentado 20190903                               Who.Validating = True
                            End If
                            oRowI.Fields("CJRMVI_NROCVT").Enabled = True
                            AsignoValor Who, oRowI.Fields("CJRMVI_NROCVT"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
                            
                            oRowI.Fields("CJRMVI_NROSVT").Enabled = True
                            AsignoValor Who, oRowI.Fields("CJRMVI_NROSVT"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
                        End If
                    End If
                    
                    If msHeadTb = "CORMVH" Or msHeadTb = "PVRMVH" Then
'                        If Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value <> "" Then
'                            Who.Validating = True
'                        End If
                        oRowI.Fields("CJRMVI_NROCPV").Enabled = True
                        oRowI.Fields("CJRMVI_NROSPV").Enabled = True
                        
                        AsignoValor Who, oRowI.Fields("CJRMVI_NROCPV"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
                        AsignoValor Who, oRowI.Fields("CJRMVI_NROSPV"), Who.Table.Rows(1).Fields(msHeadTb & "_NROCTA").Value, False
                        
                    End If
                
                End If
        
                nImporte = oRowI.Fields("CJRMVI_IMPORT").Value
                sModCom = "CJ"
                If msHeadTb = "CJRMVH" Then
                    
                    sCodCom = Who.Table.Rows(1).Fields(msHeadTb & "_CODFOR").Value
                
                Else
                    sCodCom = Who.Table.Rows(1).Fields(msHeadTb & "_CODCOM").Value
                    
                    Dim bCODFCJ, oField
                    bCODFCJ = False
                    
                    If ExisteCampo(Who.Table.Rows(1).Fields, "VIRT_CODFCJ") Then
                            bCODFCJ = True
                            sCodCom = Who.Table.Rows(1).Fields("VIRT_CODFCJ").Value
                    End If
                    
                End If
            sCirCom = ""
            sCirApl = ""
            sTipReg = Who.Table.Rows(1).Fields("USR_" & msHeadTb & "_" & sTipComReg).Value
            sNroCta = oRowI.Fields("CJRMVI_NROCVT").Value
            sTipCpt = oRowI.Fields("CJRMVI_TIPCPT").Value
            sCodCpt = oRowI.Fields("CJRMVI_CODCPT").Value
            sCuenta = oRowI.Fields("CJRMVI_CUENTA").Value
            sTipPro = ""
            sArtCod = ""
            sDeposi = ""
            sSector = ""
            sDepIng = ""
            sSecIng = ""
            sNotros = ""
            nNroLot = 0
            nNroIlo = 0
    
    
            sSql = "EXEC SP_CG_DLL_BuscoCTItem 'DLL', '" & sModCom & "','" & sCodCom & "','" & sCirCom & "','" & sCirApl & "','" & sTipReg & "','" & sNroCta & "','" & sTipCpt & "','" & sCodCpt & "','" & sCuenta & "','" & sTipPro & "','" & sArtCod & "','" & sDeposi & "','" & sSector & "','" & sDepIng & "','" & sSecIng & "', '" & sNotros & "'"
            Set Rs = Who.DataAccess.OpenResultset(CStr(sSql))
    
            
            If Not Rs.EOF Then
                   nCtaMax = Rs("CuentaApertura").Value   ' Cantidad maxima de apertura
                   sCuenta = Rs("CuentaValor").Value      ' Cuenta Contable
                   sCtaISC = Rs("CuentaImpDim").Value     ' Cuenta es imputable in dimensiones
                   sCtaEdi = Rs("CuentaEdita").Value      ' Cuenta es editable
                   nSCsMax = Rs("SubCueApertura").Value   ' cantida maxima de apertura de subcuenta por cuenta
                   sSCIEdi = "S"                          ' define si edito el importe
                   sRLsVal = Rs("RLsValor").Value         ' Valor del RL
                   iRLsItm = Rs("RLsItem").Value          ' Item de la USR_VTMRLS del Cliente
                   sRPsVal = Rs("RPsValor").Value         ' Valor del RP
                   sRLsEdi = Rs("RLsEdita").Value         ' RL es editable
                   sRPsEdi = Rs("RPsEdita").Value         ' RP es editable
                   sCodVin = ""                           ' Codigo vinulacion siempre es blanco
                   sCVsEdi = Rs("CVsEdita").Value         ' CV es editable
                   sCVsReq = Rs("CVsRequerido").Value     ' CV es Requerido
                   sCVsCnd = Rs("CVsCondicion").Value     ' CV Condicion de validacion
            Else
                Exit Function
            End If
            Rs.Close
    
    
            Erase aSubcuentas
            nIndice = 1
            If Left(sCuenta, 3) <> "CPT" And Left(sRLsVal, 3) <> "CPT" And Left(sRPsVal, 3) <> "CPT" Then
                'Me fijo y actualizo la cantidad de items
                aSubcuentas(0, 0) = nIndice
    
                'actualizo el array
                aSubcuentas(1, nIndice) = sRLsVal
                aSubcuentas(2, nIndice) = sRPsVal
                aSubcuentas(3, nIndice) = sRLsVal & sRPsVal
                aSubcuentas(4, nIndice) = nImporte
                aSubcuentas(5, nIndice) = ""
    
            Else
                If Left(sRLsVal, 3) = "CPT" Then
                    RecuperoSCSolapa Who, "RL", Right(sRLsVal, 1), sRPsVal
                End If
    
                If Left(sRPsVal, 3) = "CPT" Then
                    RecuperoSCSolapa Who, "RP", Right(sRPsVal, 1), sRLsVal
                End If
    
            End If
            
            If Trim(sCuenta) <> "" Then
                AsignoValor Who, oRowI.Fields("CJRMVI_CUENTA"), sCuenta, False
            End If
            
            If sCtaISC = "S" Then
                oRowI.Fields("CJRMVI_NROCVT").ReadOnly = False
    
    '***************************************** Tratamiento de la SubCuenta, solo si la cuenta determina que utilizo subcuenta
                    
                    If oRowI.Fields("VIRT_TABLA").ReadOnly = False Then
                        For Each oRowD In oRowI.Tables(sAliasD).Rows
                            oRowD.SetError ""
                        Next
                                   'Recorro la solapa Asiento Contable
                        'If mvEvent = 27 Or mvEvent = 28 Then  comentado 'AF 10.4
                        If mvEvent = 27 Or Who.Saving = True Or mPasaValImport Then '(mPasaValImport = True And msMainTb = "CJRMVH") Then
                            For Each oRowD In oRowI.Tables(sAliasD).Rows
                            
                            Dim BorroG As Boolean
                            BorroG = True
                            Dim oRowGx As Object
                            For Each oRowGx In oRowD.Tables(sAliasG).Rows
                                If Trim(oRowGx.Fields("CJRMVG_SUBCUE").Value) <> CStr(Trim(aSubcuentas(3, nIndice))) And Trim(oRowGx.Fields("CJRMVG_SUBCUE").Value) <> "" Then
                                    BorroG = False
                                Else
                                    If Trim(oRowGx.Fields("CJRMVG_SUBCUE").Value) = Trim(aSubcuentas(cSubCuenta, nIndice)) And _
                                       Trim(oRowGx.Fields("CJRMVG_CUENTA").Value) = Trim(sCuenta) And _
                                       Trim(oRowGx.Fields("USR_CJRMVG_CODIRL").Value) = Trim(aSubcuentas(cCodiRL, nIndice)) And _
                                       Trim(oRowGx.Fields("USR_CJRMVG_CODIRP").Value) = Trim(aSubcuentas(cCodiRP, nIndice)) And _
                                       oRowGx.Fields("CJRMVG_IMPORT").Value = aSubcuentas(cImporte, nIndice) Then
                                        BorroG = False
                                    End If
                                End If
                            Next
                            
                            If (BorroG) Then
                                If oRowD.Tables(sAliasG).Rows.Count = 0 Then
                                   oRowD.Tables(sAliasG).Rows.Add (1)
                                   oRowD.SetError ""
                                End If
                                
                                Dim sCodDis As String
                                sCodDis = aSubcuentas(3, 1)
    
                                If oRowD.Tables(sAliasG).Rows.Count > 0 Then
                                    For nIndice2 = 1 To aSubcuentas(0, 0)
                                        Dim iTest As Integer
    
                                        If nIndice2 > 1 Then
                                           oRowD.Tables(sAliasG).Rows.Add (1)
                                        End If
                                        
                                        With oRowD.Tables(sAliasG).Rows(oRowD.Tables(sAliasG).Rows.Count)
    
                                            AsignoValor Who, .Fields("CJRMVG_CUENTA"), Trim(sCuenta), .Fields("CJRMVG_CUENTA").ReadOnly
                                            AsignoValor Who, .Fields("CJRMVG_CODDIM"), "RLRP", .Fields("CJRMVG_CODDIM").ReadOnly
                                            AsignoValor Who, .Fields("CJRMVG_DIMSUB"), "RLRP", .Fields("CJRMVG_DIMSUB").ReadOnly
                                            
                                            If aSubcuentas(cCodiRL, nIndice) <> "" Then
                                                AsignoValor Who, .Fields("USR_CJRMVG_CODIRL"), Trim(aSubcuentas(cCodiRL, nIndice)), .Fields("USR_CJRMVG_CODIRL").ReadOnly
                                            End If
                                            If aSubcuentas(cCodiRP, nIndice) <> "" Then
                                                AsignoValor Who, .Fields("USR_CJRMVG_CODIRP"), Trim(aSubcuentas(cCodiRP, nIndice)), .Fields("USR_CJRMVG_CODIRP").ReadOnly
                                            End If
                                            If aSubcuentas(cSubCuenta, nIndice) <> "" Then
                                                AsignoValor Who, .Fields("CJRMVG_SUBCUE"), Trim(aSubcuentas(cSubCuenta, nIndice)), .Fields("CJRMVG_SUBCUE").ReadOnly
                                            End If
                                            
                                            AsignoValor Who, .Fields("CJRMVG_IMPORT"), Trim(aSubcuentas(cImporte, nIndice)), False
                                        
                                        End With
                                    
                                    Next
                                End If
    
                                If mvEvent = 27 Or Who.Saving Then
                                    
                                    AsignoValor Who, oRowD.Fields("CJRMVD_CODDIM"), "RLRP", False
                                    AsignoValor Who, oRowD.Fields("CJRMVD_IMPUTA"), "S", False
                                    AsignoValor Who, oRowD.Fields("CJRMVD_TOTIMP"), oRowI.Fields("CJRMVI_IMPORT").Value, oRowD.Fields("CJRMVD_TOTIMP").ReadOnly
                                
                                End If
                            End If
                        Next
                    Else 'Si está grisado el jump de subcuentas
                        
                        
                    End If
                End If
            End If
            
            If sCtaEdi = "N" Then
                oRowI.Fields("CJRMVI_CUENTA").ReadOnly = True
            Else
                oRowI.Fields("CJRMVI_CUENTA").ReadOnly = False
            End If
            
            If msHeadTb = "CJRMVH" Then
                '--------------------------------------Busco SN--------------------------------------------------
                If sCodigoSN = vbNullString Then
                    '@ModCom CHAR(2) ,@CodCom VARCHAR(6),@CirCom VARCHAR(6),@CirApl VARCHAR(6),@TipReg VARCHAR(6),@Nrocta VARCHAR(25)
                    sSQL2 = "EXEC SP_CG_DLL_BuscoSN '" & sModCom & "','" & sCodCom & "','" & sCirCom & "','" & sCirApl & "','" & sTipReg & "','" & sNroCta & "','" & sTipCpt & "'"
                    Set Rs2 = Who.DataAccess.OpenResultset(CStr(sSQL2))
                    If Not Rs2.EOF Then
                        sCodigoSN = ""
                        sCodigoSN = Rs2("NuevoCodigoSN").Value   ' Código SN
                        sDescSN = Rs2("DescripcionSN").Value
                    End If
        
                    If Trim(sCodigoSN) <> "" Then
                        With Who.Table.Rows(1)
                            AsignoValor Who, .Fields("USR_" & msHeadTb & "_CODISN"), Trim(sCodigoSN), True
                            AsignoValor Who, .Fields("USR_" & msHeadTb & "_DESCSN"), Trim(sDescSN), True
                        End With
                    End If
                End If
            End If
        End If
    Next
    
    Set Rs = Nothing
    Set Rs2 = Nothing   'AF 10.4
    End If
End Function


Private Function CompletaNROCVT(Who As Object, sAlias As String)
    Dim oRowI As Object
    Dim RsCVT As rdoResultset
    Dim sSQLCVT As String
        
    Who.IntoEvent = False
    
    With Who.Table.Rows(1).Tables(sAlias)
        For Each oRowI In .Rows
                With oRowI
                    If Trim(oRowI.Fields("CJRMVI_NROCVT").Value) <> "" Then
                        sSQLCVT = "SELECT ISNULL(VTMCLH_NOMBRE,'') VTMCLH_NOMBRE FROM VTMCLH WHERE VTMCLH_NROCTA = '" & Trim(oRowI.Fields("CJRMVI_NROCVT").Value) & "'"
                        Set RsCVT = Who.DataAccess.OpenResultset(CStr(sSQLCVT))
                        If Not RsCVT.EOF Then
                            AsignoValor Who, oRowI.Fields("CJRMVI_NOMBVT"), RsCVT("VTMCLH_NOMBRE").Value, oRowI.Fields("CJRMVI_NOMBVT").ReadOnly
                       End If
                    End If
                End With
        Next
    End With
End Function


Private Function CompletaNROCPV(Who As Object, sAlias As String)
    Dim oRowI As Object
    Dim RsCPV As rdoResultset
    Dim sSQLCPV As String
        
    Who.IntoEvent = False
    With Who.Table.Rows(1).Tables(sAlias)
        For Each oRowI In .Rows
                With oRowI
                    If Trim(oRowI.Fields("CJRMVI_NROCPV").Value) <> "" Then
                        sSQLCPV = "SELECT ISNULL(PVMPRH_NOMBRE,'') PVMPRH_NOMBRE FROM PVMPRH WHERE PVMPRH_NROCTA = '" & Trim(oRowI.Fields("CJRMVI_NROCPV").Value) & "'"
                        Set RsCPV = Who.DataAccess.OpenResultset(CStr(sSQLCPV))
                        If Not RsCPV.EOF Then
                            AsignoValor Who, oRowI.Fields("CJRMVI_NOMBPV"), RsCPV("PVMPRH_NOMBRE").Value, oRowI.Fields("CJRMVI_NOMBPV").ReadOnly
                       End If
                    End If
                End With
        Next
    End With

End Function


Private Function RecuperaImporteNetoPagoCob(Who As Object, sAlias As String, iNroItm As Long) As Double
Dim dImporte As Double
Dim oRowI As Object
Dim sAliasItems As String
sAliasItems = ""
dImporte = 0

If iNroItm = 1 Then
    Select Case sAlias
        Case "PVRRCC04"
              sAliasItems = "PVRRCC01"
        Case "PVRRCC05"
              sAliasItems = "PVRRCC01"
        Case "VTRRCC04"
              sAliasItems = "VTRRCC01"
        Case "VTRRCC05"
              sAliasItems = "VTRRCC01"
    End Select
    If sAliasItems <> "" Then
        With Who.Table.Rows(1).Tables(sAliasItems)
            For Each oRowI In .Rows
                dImporte = dImporte + CDbl(oRowI.Fields("VIRT_NETO").Value)
            Next
        End With
    End If
End If
RecuperaImporteNetoPagoCob = dImporte
End Function

'*******************************************
'                        Grabo RL - RP - CV  en el comprobante generado
Private Sub Grabo_RL_RP_VIN(Who As Object, sAlias As String, sAliasD As String, sAliasG As String)
'*******************************************
    Dim Rs
    Dim sSql
    Dim sModFor
    Dim sCodFor
    Dim nNroFor
    Dim nNroitm
    Dim nNroitd
    Dim nNroits
    Dim sCodRLs
    Dim sCodRPs
    Dim sCodVin
    Dim nIndice
    Dim sDebHab
    Dim sCodDim As String
    Dim sSubcue As String
    
    Dim oRowI
    Dim oRowD
    Dim oRowG
    
    'Como en la G no puedo grabar un campo normal, lo hago por aca con los de otro tipo
    For Each oRowI In Who.Table.Rows(1).Tables(sAlias).Rows
        For Each oRowD In oRowI.Tables(sAliasD).Rows
            For Each oRowG In oRowD.Tables(sAliasG).Rows
'                For nIndice = 1 To Who.DataAccess.PerformedOperations.Count
                    
                    sModFor = oRowG.Fields("CJRMVG_MODFOR").Value 'Who.DataAccess.PerformedOperations(nIndice).Keys(1).Value
                    sCodFor = oRowG.Fields("CJRMVG_CODFOR").Value 'Who.DataAccess.PerformedOperations(nIndice).Keys(2).Value
                    nNroFor = oRowG.Fields("CJRMVG_NROFOR").Value 'Who.DataAccess.PerformedOperations(nIndice).Keys(3).Value
                    nNroitm = oRowG.Fields("CJRMVG_NROITM").Value
                    nNroitd = oRowG.Fields("CJRMVG_NROITD").Value
                    sDebHab = oRowG.Fields("CJRMVG_DEBHAB").Value
                    nNroits = oRowG.Fields("CJRMVG_NROITS").Value
                    sCodDim = oRowG.Fields("CJRMVG_CODDIM").Value
                    sSubcue = oRowG.Fields("CJRMVG_SUBCUE").Value
                    
                    sCodRLs = oRowG.Fields("USR_CJRMVG_CODIRL").Value
                    sCodRPs = oRowG.Fields("USR_CJRMVG_CODIRP").Value
                    sCodVin = oRowG.Fields("USR_CJRMVG_CODVIN").Value
    
                    sSql = " "
                    sSql = sSql & " UPDATE CJRMVG  "
                    sSql = sSql & "    SET USR_CJRMVG_AUXIRL = '" & sCodRLs & "', "
                    sSql = sSql & "        USR_CJRMVG_AUXIRP = '" & sCodRPs & "', "
                    sSql = sSql & "        USR_CJRMVG_AUXVIN = '" & sCodVin & "' "
                    sSql = sSql & " WHERE CJRMVG_MODFOR  = '" & sModFor & "'"
                    sSql = sSql & "   AND CJRMVG_CODFOR  = '" & sCodFor & "'"
                    sSql = sSql & "   AND CJRMVG_NROFOR  = '" & nNroFor & "'"
                    sSql = sSql & "   AND CJRMVG_NROITM  = '" & nNroitm & "'"
                    sSql = sSql & "   AND CJRMVG_NROITD  = '" & nNroitd & "'"
                    sSql = sSql & "   AND CJRMVG_DEBHAB  = '" & sDebHab & "'"
                    sSql = sSql & "   AND CJRMVG_NROITS  = '" & nNroits & "'"
                    sSql = sSql & "   AND CJRMVG_CODDIM  = '" & sCodDim & "'"
                    sSql = sSql & "   AND CJRMVG_SUBCUE  = '" & sSubcue & "'"

                    sSql = sSql & " ; "

                    Set Rs = Who.DataAccess.OpenResultset(sSql)
                    Rs.Close
                
'                Next
            Next
        Next
    Next

End Sub
Private Function ValidaImportDConI(Who As Object) As Integer
    Dim Valida As Integer
    Dim oRow1 As Object
    Dim oRowD1 As Object
    Dim sCurTabDH As String
    Dim sCurTabDHD As String
    Dim sCurTabDHG As String
    sCurTabDH = vbNullString
    If Who.Saving Then
        Exit Function
    End If
    Select Case Who.CurrentTable
        Case msDebTbG
            sCurTabDH = msDebTb
            sCurTabDHD = msDebTbD
            sCurTabDHG = msDebTbG
        Case msHabTbG
            sCurTabDH = msHabTb
            sCurTabDHD = msHabTbD
            sCurTabDHG = msHabTbG
    End Select
    
    If Trim(sCurTabDHG) <> "" And msHeadTb = "CJRMVH" Then
        With Who.Table.Rows(1).Tables(sCurTabDH)
               Set oRow1 = .Rows(.CurrentRow)
               With oRow1
               
                 If .Fields("CJRMVI_IMPORT").Value <> 0 Then
                    Set oRowD1 = .Tables(sCurTabDHD).Rows(1)
                    With oRowD1
                        oRowD1.SetError ""
                        If oRow1.Fields("CJRMVI_IMPORT").Value <> .Fields("CJRMVD_TOTIMP").Value Then
                            oRow1.Fields("CJRMVI_IMPORT").Value = .Fields("CJRMVD_TOTIMP").Value
                        End If
                        oRowD1.Fields("CJRMVD_IMPUTA").Value = "S"
                    End With
                 End If
          
               End With
        End With
   End If
End Function

Private Sub AsignoValor(Who As Object, oField As Object, vValor As Variant, bReadOnly As Boolean)
Dim sFName
sFName = oField.Name
    If oField.Value <> vValor Then
        
        If oField.ColumnDef.Validate = True Then
            Who.Validating = True
        End If
        oField.Enabled = True
        oField.Value = vValor
        oField.ReadOnly = bReadOnly
        
    End If

End Sub

Private Function ExisteCampo(oFields As Object, sField As String) As Boolean

    Dim oField As Object

    For Each oField In oFields
    
        If oField.Name = sField Then
        
            ExisteCampo = True
            Exit For
        
        End If
    
    Next

End Function

