Option Explicit On

Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Xml
Imports Microsoft.VisualBasic
Imports Windows.Win32.System
Imports System.Globalization

Module Module1



    Public sop As Integer  'start of period
    Public eop As Integer  'end of period
    Public timestep As Integer ' timestep-length in years
    Public Const speedImport As Boolean = True


    ' ATTENTION: DO not use double multipliation in ACCESS queries: 3*0.001*SUM(). This will result in wrong results


    '1. Total Primary Energy Supply by Type
    Function SQLquery1(caseName As String, region As String) As String

        Const supplyBio As String = "('BDW', 'BOC', 'BST', 'BSU', 'BWO', 'BBDX', 'BETX', 'BNGX')"
        Const supplyOil As String = "('DSX', 'GSX', 'OIX', 'OI1', 'OI2', 'OI3', 'OI4', 'OI5')"
        Const supplySyn As String = "('HGX', 'FGX', 'FHX', 'BMTX')"
        Const supplyGas As String = "('LNG', 'NGI', 'NGG', 'NG1', 'NG2', 'NG3', 'NG4', 'NG5', 'NG6')"
        Const supplyCoal As String = "('BCO', 'COA', 'HCX')"

        Const supFos As String = "('SUPFOS', 'R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')"
        'Const renElcWithoutHydro As String = "('E41', 'E42', 'E61', 'E62', 'E81', 'H41', 'H42', 'H61', 'S1G')"
        'Const renHeat As String = "('S1G', 'E75')"

        Dim s1 As String
        Dim sRen As String
        Dim sRenWithoutHydro As String
        Dim sHydro As String
        Dim sNonElcRen As String
        Dim sNuclear As String
        Dim sElcTrade As String
        Dim sGas As String

        s1 =
        " SELECT 'Biomass' AS Energy, " & template("0.001*SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-1,1)*T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter IN " & supFos & " AND Region IN (" & region & ") " &
        " AND Item2 IN " & supplyBio & " AND Item2 NOT LIKE 'ETL%' UNION " &
        " SELECT 'Oil' AS Energy, " & template("0.001*SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-1,1)*T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter IN " & supFos & " AND Region IN (" & region & ") " &
        " AND Item2 IN " & supplyOil & " AND Item2 NOT LIKE 'ETL%' UNION " &
        " SELECT 'Coal' AS Energy, " & template("0.001*SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-1,1)*T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter IN " & supFos & " AND Region IN (" & region & ") " &
        " AND Item2 IN " & supplyCoal & " AND Item2 NOT LIKE 'ETL%' " & " UNION " &
        " SELECT 'Synfuel (Trade)' AS Energy, " & template("0.001*SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-1,1)*T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter IN " & supFos & " AND Region IN (" & region & ") " &
        " AND Item2 IN " & supplySyn & " AND Item2 NOT LIKE 'ETL%' "

        sRen = " SELECT 'Renewables' AS Energy, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter  = 'TOT.USEREN' AND Region IN (" & region & ") "

        sRenWithoutHydro =
        " SELECT Energy, " & template("t1.TT2000 - IIF(ISNULL(t2.TT2000),0.00,t2.TT2000) AS 2000", sop, timestep, eop) &
        " FROM (" &
        "   (SELECT 'dummy' AS Dummy, 'Renewables' AS Energy, " & template("0.001*SUM(T2000) AS TT2000", sop, timestep, eop) &
        "    FROM tblTRESULTS " &
        "    WHERE CaseName = '" & caseName & "' AND Parameter  = 'TOT.USEREN' " &
        "       AND Region IN (" & region & ") " &
        "    ) AS t1 " &
        "  INNER JOIN " &
        "    (SELECT 'dummy' AS Dummy, " & template("0.001*SUM(T2000) AS TT2000", sop, timestep, eop) &
        "     FROM tblTRESULTS " &
        "     WHERE CaseName = '" & caseName & "' AND Parameter  = 'OUTPUT.ELC' " &
        "        AND Item1 = 'E31' AND Region IN (" & region & ") " &
        "     ) AS t2 " &
        "  ON t1.Dummy = t2.Dummy  ) "


        sNonElcRen =
        " SELECT 'Solar Fuels' AS Energy, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter  = 'OUTOTH.PRC.A??' " &
        " AND Item1 IN ('F42','H41','H41G') AND Region IN (" & region & ") "

        sHydro =
        " SELECT 'Hydro' AS Energy, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter  = 'OUTPUT.ELC' " &
        " AND Item1 = 'E31' AND Region IN (" & region & ") "


        sNuclear =
        " SELECT 'Nuclear' AS Energy, " & template("0.003*SUM(T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter  = 'OUTPUT.ELC' " &
        " AND Item1 IN ('E21','E22') AND Region IN (" & region & ") "

        sElcTrade =
        " SELECT 'Electricity' AS Energy, " & template("0.001*SUM(IIF(Item1 LIKE 'EXP%',-1,1)*T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Parameter = 'R_TSEPELC.L' AND Region IN (" & region & ") "

        'sGas = _
        '" SELECT 'Gas' AS Energy, " & template("0.001*SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L' OR (Parameter = 'R_TSEP.L' AND Item1 LIKE 'EXP%'),-1,1)*T2000) AS 2000", sop, timestep, eop) & _
        '" FROM tblTRESULTS " & _
        '" WHERE CaseName = '" & caseName & "' AND Region IN (" & region & ") AND (" & _
        '" (Parameter IN " & supFos & " AND Item2 IN " & supplyGas & " AND Item2 NOT LIKE 'ETL%' ) OR " & _
        '"  Parameter = 'R_TSEP.L' ) "

        sGas =
        " SELECT 'Gas' AS Energy, " & template("0.001*SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L' ,-1,1)*T2000) AS 2000", sop, timestep, eop) &
        " FROM tblTRESULTS " &
        " WHERE CaseName = '" & caseName & "' AND Region IN (" & region & ") AND (" &
        " (Parameter IN " & supFos & " AND Item2 IN " & supplyGas & " AND Item2 NOT LIKE 'ETL%' ) " &
        "  ) "


        SQLquery1 = s1 & " UNION " & sHydro &
                 " UNION " & sRenWithoutHydro &
                 " UNION " & sNuclear &
                 " UNION " & sElcTrade &
                 " UNION " & sGas &
                 " UNION " & sNonElcRen

    End Function


    '2. Total Primary Energy Supply by Region
    Function SQLquery2(caseName As String) As String

        Const supplyBio As String = "'BDW', 'BOC', 'BST', 'BSU', 'BWO', 'BBDX', 'BETX', 'BNGX'"
        Const supplyOil As String = "'DSX', 'GSX', 'OIX', 'OI1', 'OI2', 'OI3', 'OI4', 'OI5'"
        Const supplySyn As String = "'HGX', 'FGX', 'FHX', 'BMTX'"
        Const supplyGas As String = "'LNG', 'NGI', 'NGG', 'NG1', 'NG2', 'NG3', 'NG4', 'NG5', 'NG6'"
        Const supplyCoal As String = "'BCO', 'COA', 'HCX'"

        Const supFos As String = "('SUPFOS', 'R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')"

        SQLquery2 =
" SELECT Region, " & template("0.001*SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L' OR (Parameter='R_TSEPELC.L' AND Item1 LIKE 'EXP%'),-1, IIF(Parameter = 'OUTPUT.ELC',3,1))*T2000) AS 2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE CaseName = '" & caseName & "'" &
"   AND Region <> 'REGION0' " &
"   AND( (Parameter IN " & supFos &
"           AND Item2 IN (" & supplyBio & " , " & supplyOil & " , " & supplyGas & " , " & supplyCoal & " , " & supplySyn & ") AND Item2 NOT LIKE 'ETL%' )" &
"        OR ( Parameter  = 'OUTPUT.ELC' AND Item1 IN ('E21','E22') ) " &
"        OR ( Parameter  = 'OUTOTH.PRC.A??' AND Item1 IN ('F42','H41','H41G') ) " &
"        OR Parameter IN ('TOT.USEREN', 'R_TSEPELC.L') " &
"       ) " &
" GROUP BY Region"

    End Function


    '3. CO2-Emissions
    Function SQLquery3(caseName As String) As String

        SQLquery3 =
 " SELECT Region, " & template("44/12*0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
 " FROM tblTRESULTS                     " &
 " WHERE CaseName  = '" & caseName & "' " &
 "   AND Parameter = 'EMISSION.L'       " &
 "   AND Item2     = 'CO2'              " &
 " GROUP BY Region                      "

    End Function

    '4. CCS
    Function SQLquery4(caseName As String) As String

        SQLquery4 =
 " SELECT Region, " & template("44 / 12 * 0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
 " FROM tblTRESULTS                     " &
 " WHERE CaseName  = '" & caseName & "' " &
 "   AND Parameter = 'EMISSION.L'       " &
 "   AND Item2     = 'COS'              " &
 " GROUP BY Region                      "

    End Function

    '5. Fuels for Long-Range Cars
    Function sqlQuery5(caseName As String, region As String)

        sqlQuery5 = fuelQuery(caseName, region, "'T2'")

    End Function

    '6. Fuels for Short-Range Cars
    Function sqlQuery6(caseName As String, region As String)

        sqlQuery6 = fuelQuery(caseName, region, "'T4'")

    End Function

    '7. Fuels for Cars
    Function sqlQuery7(caseName As String, region As String)

        sqlQuery7 = fuelQuery(caseName, region, "'T2', 'T4'")

    End Function

    ' 8. Fuels in Transport
    Function SQLquery8(caseName As String, region As String) As String

        SQLquery8 =
     "SELECT Technology, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 IN ('GLR', 'DLR', 'OPR', 'AVT'),               'Oil Type',      " &
     "  IIF(Item2 IN ('ALR'),                                    'Alc in other surf', " &
     "  IIF(Item2 IN ('HDN'),                                    'Hydrogen',      " &
     "  IIF(Item2 IN ('ELC'),                                    'Electricity',   " &
     "  IIF(Item2 IN ('CGR', 'CGR1'),                            'Gas Type',      " &
     "  'Other (Coal)'))))) AS Technology           " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('T1','T2', 'T3', 'T4') " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Technology                 "

    End Function


    '9. Technology Mix for Long-Range Cars
    Function SQLquery9(caseName As String, region As String) As String

        SQLquery9 =
" SELECT Technology, " & template("SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM (SELECT *, " &
"  IIF(Item1 IN ( 'T2GI', 'T2GA', 'T2DI', 'T2DA'), 'Liquid Fuel ICEV', " &
"  IIF(Item1 IN ('T2GH', 'T2DH'),                  'Liquid Fuel Hybrid', " &
"  IIF(Item1 = 'T2EH',                             'Liquid Fuel Plug-in', " &
"  IIF(Item1 = 'T2CI',                             'Gas Fuel ICEV', " &
"  IIF(Item1 = 'T2CH',                             'Gas Fuel Hybrid', " &
"  IIF(Item1 = 'T2HH',                             'Hydrogen Hybrid', " &
"  IIF(Item1 = 'T2GF',                             'Gasoline Fuel Cell', " &
"  IIF(Item1 = 'T2HF',                             'Hydrogen Fuel Cell', " &
"  IIF(Item1 = 'T2EB',          'Electric Vehicle'  , Item1))))))))) AS Technology " &
" FROM tblTRESULTS " &
" WHERE CaseName  = '" & caseName & "' " &
"   AND Parameter = 'USENRG.TRN_DMD_DM' " &
"   AND Item2 = 'T2' " &
"   AND Region IN (" & region & ") " &
") GROUP BY Technology"

    End Function

    '10. Technology Mix of Long-Range Cars (Detailed)
    Function sqlQuery10(caseName As String, region As String) As String

        sqlQuery10 =
" SELECT Technology, " & template("SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM " &
"( SELECT *, Item1 AS Technology         " &
"  FROM tblTRESULTS                      " &
"  WHERE CaseName  = '" & caseName & "'  " &
"   AND Parameter = 'USENRG.TRN_DMD_DM'  " &
"   AND Item2 = 'T2'                     " &
"   AND Region IN (" & region & ")       " &
" ) GROUP BY Technology                   "

    End Function

    '11. Technology Mix for Short-Range Cars
    Function SQLquery11(caseName As String, region As String) As String

        SQLquery11 =
" SELECT Technology, " & template("SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM (SELECT *, " &
"  IIF(Item1 IN ('T4GI'), 'Gasoline ICEV City',     " &
"  IIF(Item1 IN ('T4GA'), 'Gasoline Adv.ICEV City', " &
"  IIF(Item1 IN ('T4GH'), 'Gasoline Hybrid City',   " &
"  IIF(Item1 IN ('T4HF'), 'Hydrogen Fuel Cell City', " &
"  IIF(Item1 IN ('T4EH'), 'Plug-in Hybrid City',     " &
"  IIF(Item1 IN ('T4EB'), 'Electric Vehicle City', Item1)))))) AS Technology " &
" FROM tblTRESULTS " &
" WHERE CaseName  = '" & caseName & "' " &
"   AND Parameter = 'USENRG.TRN_DMD_DM' " &
"   AND Item2 = 'T4' " &
"   AND Region IN (" & region & ") " &
") GROUP BY Technology"

    End Function

    '12. Technology Mix for all Cars
    Function SQLquery12(caseName As String, region As String) As String

        SQLquery12 =
" SELECT Technology, " & template("SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM (SELECT *, " &
"  IIF(Item1 IN ('T2GI', 'T2GA', 'T2DI', 'T2DA', 'T4GI', 'T4GA'), 'Liquid Fuel ICEV', " &
"  IIF(Item1 IN ('T2GH', 'T2DH', 'T4GH'), 'Liquid Fuel Hybrid',  " &
"  IIF(Item1 IN ('T2EH', 'T4EH'),         'Liquid Fuel Plug-in', " &
"  IIF(Item1 = 'T2CI',                    'Gas Fuel ICEV',       " &
"  IIF(Item1 = 'T2CH',                    'Gas Fuel Hybrid',     " &
"  IIF(Item1 = 'T2HH',                    'Hydrogen Hybrid',     " &
"  IIF(Item1 = 'T2GF',                    'Gasoline Fuel Cell',  " &
"  IIF(Item1 IN ('T2HF','T4HF'),          'Hydrogen Fuel Cell',  " &
"  IIF(Item1 IN ('T2EB','T4EB'),          'Electric Vehicle'  , Item1))))))))) AS Technology " &
" FROM tblTRESULTS " &
" WHERE CaseName  = '" & caseName & "' " &
"   AND Parameter = 'USENRG.TRN_DMD_DM' " &
"   AND Item2 IN ('T2','T4') " &
"   AND Region IN (" & region & ") " &
") GROUP BY Technology"


    End Function
    '13. New Long-Range Cars
    Function SQLquery13(caseName As String, region As String, numCarScale As Double) As String

        SQLquery13 =
"SELECT * FROM       " &
"(SELECT Technology, " & template("$*SUM(T2000) AS 2000", sop, timestep, eop, 1 / numCarScale) &
" FROM (SELECT " &
template("T2000 * " &
" IIF(Region IN ('AUSNZL'),1/135, IIF(Region IN ('ASIAPAC','LAC','BRAZIL'),1/156, IIF(Region IN ('CANMEX'),1/142, " &
" IIF(Region IN ('CENASIA','RUSSIA'),1/132, IIF(Region IN ('CHINAREG'),1/152, IIF(Region IN ('EEUR','SSAFRICA'),1/124, " &
" IIF(Region IN ('INDIA'),1/147, IIF(Region IN ('JPKRTW'),1/85," &
" IIF(Region IN ('USA','MENA'),1/182,1/140))))))))) AS T2000", sop, timestep, eop) & ", " &
"  IIF(Item1 IN ( 'T2GI', 'T2GA', 'T2DI', 'T2DA'), 'Liquid Fuel ICEV', " &
"  IIF(Item1 IN ('T2GH', 'T2DH'), 'Liquid Fuel Hybrid',   " &
"  IIF(Item1 IN ('T2EH'),         'Liquid Fuel Plug-in',  " &
"  IIF(Item1 IN ('T2CI'),         'Gas Fuel ICEV',        " &
"  IIF(Item1 IN ('T2CH'),         'Gas Fuel Hybrid',      " &
"  IIF(Item1 IN ('T2HH'),         'Hydrogen Hybrid',      " &
"  IIF(Item1 IN ('T2GF'),         'Gasoline Fuel Cell',   " &
"  IIF(Item1 IN ('T2HF'),         'Hydrogen Fuel Cell',   " &
"  IIF(Item1 IN ('T2EB'),         'Electric Vehicle','delete'))))))))) AS Technology " &
" FROM tblTRESULTS                     " &
" WHERE CaseName  = '" & caseName & "' " &
"   AND Parameter = 'INVEST.L'         " &
"   AND Region IN (" & region & ")     " &
") GROUP BY Technology) WHERE Technology <> 'delete'"

    End Function

    '14. New Short-Range Cars
    Function SQLquery14(caseName As String, region As String, numCarScale As Double) As String

        SQLquery14 =
"SELECT * FROM       " &
"(SELECT Technology, " & template("$*SUM(T2000) AS 2000", sop, timestep, eop, 1 / numCarScale) &
" FROM (SELECT " & template("1/150*T2000 AS T2000", sop, timestep, eop) & ", " &
"  IIF(Item1 IN ('T4GI'), 'Gasoline ICEV City',               " &
"  IIF(Item1 IN ('T4GA'), 'Gasoline Adv.ICEV City',           " &
"  IIF(Item1 IN ('T4GH'), 'Gasoline Hybrid City',             " &
"  IIF(Item1 IN ('T4HF'), 'Hydrogen Fuel Cell City',           " &
"  IIF(Item1 IN ('T4EH'), 'Plug-in Hybrid City',               " &
"  IIF(Item1 IN ('T4EB'), 'Electric Vehicle City', 'delete')))))) AS Technology " &
" FROM tblTRESULTS                     " &
" WHERE CaseName  = '" & caseName & "' " &
"   AND Parameter = 'INVEST.L'         " &
"   AND Region IN (" & region & ")     " &
") GROUP BY Technology) WHERE Technology <> 'delete'"

    End Function


    '15. New Cars
    Function SQLquery15(caseName As String, region As String, numCarScale As Double) As String

        SQLquery15 =
"SELECT * FROM       " &
"(SELECT Technology, " & template("$*SUM(T2000) AS 2000", sop, timestep, eop, 1 / numCarScale) &
" FROM (SELECT " &
template("T2000 * " &
" IIF(Region IN ('AUSNZL'),1/135, IIF(Region IN ('ASIAPAC','LAC','BRAZIL'),1/156, IIF(Region IN ('CANMEX'),1/142, " &
" IIF(Region IN ('CENASIA','RUSSIA'),1/132, IIF(Region IN ('CHINAREG'),1/152, IIF(Region IN ('EEUR','SSAFRICA'),1/124, " &
" IIF(Region IN ('INDIA'),1/147, IIF(Region IN ('JPKRTW'),1/85," &
" IIF(Region IN ('USA','MENA'),1/182,1/140))))))))) AS T2000", sop, timestep, eop) & ", " &
"  IIF(Item1 IN ('T2GI', 'T2GA', 'T2DI', 'T2DA', 'T4GI', 'T4GA'), 'Liquid Fuel ICEV', " &
"  IIF(Item1 IN ('T2GH', 'T2DH', 'T4GH'), 'Liquid Fuel Hybrid',   " &
"  IIF(Item1 IN ('T2EH', 'T4EH'),         'Liquid Fuel Plug-in',  " &
"  IIF(Item1 IN ('T2CI'),                 'Gas Fuel ICEV',        " &
"  IIF(Item1 IN ('T2CH'),                 'Gas Fuel Hybrid',      " &
"  IIF(Item1 IN ('T2HH'),                 'Hydrogen Hybrid',      " &
"  IIF(Item1 IN ('T2GF'),                 'Gasoline Fuel Cell',   " &
"  IIF(Item1 IN ('T2HF', 'T4HF'),         'Hydrogen Fuel Cell',   " &
"  IIF(Item1 IN ('T2EB', 'T4EB'),         'Electric Vehicle','delete'))))))))) AS Technology " &
" FROM tblTRESULTS                     " &
" WHERE CaseName  = '" & caseName & "' " &
"   AND Parameter = 'INVEST.L'         " &
"   AND Region IN (" & region & ")     " &
") GROUP BY Technology) WHERE Technology <> 'delete'"

    End Function

    '16. Fuel Refining and Synthesis
    Function sqlQuery16(caseName As String, region As String) As String

        sqlQuery16 =
"SELECT * FROM " &
"(SELECT Technology, " & template(" 0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM (SELECT *, " &
"   IIF(Item1 IN ('BB1', 'BC1', 'BS1', 'BWA', 'BW1', 'BW2', 'BW3', 'BW4', 'BW5', 'BO1'), 'Biofuels', " &
"   IIF(Item1 IN ('H01', 'H01S', 'H01SF', 'H11', 'H11S', 'H21HP', 'H21HT', 'H21SI', 'H41', 'H42', 'H61', " &
"                 'H71', 'H72', 'H73', 'H80', 'H90', 'H91', 'HBD11', " &
"                 'HSD11', 'HSD70', 'HSD85', 'HSD90'), 'Hydrogen',     " &
"   IIF(Item1 = 'SRN',           'Coal-to-Liquid',                     " &
"   IIF(Item1 IN('NTF','NAF','F41'),   'Gas-to-Liquid',                " &
"   IIF(Item1 = 'HTF',           'Hydrogen-to-Liquid',                 " &
"   IIF(Item1 IN ('BG1', 'BG2', 'BG3'), 'Green Bio-to-Liquid', " &
"   IIF(Item1 IN ('HFG'),               'Green H2-to-Liquid',  " &
"   IIF(Item1 IN ('F42'),               'Green H2O-to-Liquid', " &
"   IIF(Item1 IN ('H41G','H61G', 'H80G', 'H90G', 'H91G', 'HBD90'), 'Green Hydrogen', " &
"   IIF(Item1 IN ('HNGAHP', 'HNGADL', 'HNGAAA'), 'Hydrogen-to-Gas',    " &
"   IIF(Item1 IN ('S3M'),        'Gas-to-MeOH',                        " &
"   IIF(Item1 IN ('S4M'),        'H2-to-MeOH',                         " &
"   IIF(Item1 IN ('SR9', 'SRA'), 'Oil products',                       " &
"   IIF(Item1 = 'SRM', 'Coal-to-MeOH', 'delete')))))))))))))) AS Technology " &
" FROM tblTRESULTS " &
"  WHERE CaseName = '" & caseName & "' " &
"    AND Parameter IN ('OUTOTH.PRC.LIQ', 'OUTOTH.PRC.GAS', 'OUTOTH.PRC.A??', 'OUTOTH.CON.TOT') " &
"    AND Region IN (" & region & ") " &
" ) GROUP BY Technology) WHERE Technology <> 'delete' "

    End Function

    '17. Biofuel Production
    Function SQLquery17(caseName As String, region As String) As String

        Dim s1 As String
        Dim s2 As String

        s1 =
"SELECT * FROM " &
"(SELECT Technology, " & template(" 0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM (SELECT *, " &
"   IIF(Item1 = 'BB1',          'Ethanol from Cellulose/Stover',   " &
"   IIF(Item1 = 'BC1',          'Ethanol from Corn',               " &
"   IIF(Item1 = 'BS1',          'Ethanol from Sugar Crops',        " &
"   IIF(Item1 = 'BW5',          'Methanol from Wood',              " &
"   IIF(Item1 = 'BO1',          'Biofuel (Diesel) from Oil Crops', " &
"   IIF(Item1 IN ('BW1','BW2'), 'Biodiesel (Green) from Wood',      " &
"   IIF(Item1 IN ('BG1','BG2','BG3'), 'Biofuel (Green) from Wood/Stover'," &
"   IIF(Item1 = 'BWA',          'SNG from Anaerobic Waste',        " &
"   IIF(Item1 IN ('BW3','BW4'), 'SNG from Wood', 'delete'))))))))) AS Technology " &
" FROM tblTRESULTS " &
"  WHERE CaseName = '" & caseName & "' " &
"    AND Parameter IN ('OUTOTH.PRC.LIQ', 'OUTOTH.PRC.GAS', 'OUTOTH.PRC.A??', 'OUTOTH.CON.TOT') " &
"    AND Region IN (" & region & ") " &
" ) GROUP BY Technology) WHERE Technology <> 'delete' "

        s2 = " SELECT 'special' AS Technology, " & template("0.0*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM tblTRESULTS " &
     "  WHERE CaseName = '" & caseName & "' " &
     "    AND Parameter IN ('OUTOTH.PRC.LIQ', 'OUTOTH.PRC.GAS', 'OUTOTH.PRC.A??', 'OUTOTH.CON.TOT') " &
     "    AND Region IN (" & region & ") "


        SQLquery17 = s1 & " UNION " & s2




    End Function

    '18. Hydrogen Production (simplifed)
    Function SQLquery18(caseName As String, region As String) As String

        Dim s As String

        s = " AS [Primary Energy], " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
        "   FROM tblTRESULTS " &
        "   WHERE CaseName = '" & caseName & "'" &
        "    AND Parameter IN ('OUTOTH.PRC.LIQ','OUTOTH.PRC.GAS','OUTOTH.PRC.A??','OUTOTH.CON.TOT')" &
        "    AND Region IN (" & region & ") " &
        "    AND Item1 IN "

        SQLquery18 =
        " SELECT 'Coal Gasification' " & s & "('H01') UNION " &
        " SELECT 'Coal Gasif. with CCS' " & s & "('H01S','H01SF') UNION " &
        " SELECT 'Natural Gas Reforming' " & s & " ('H11', 'HBD11','HSD11') UNION " &
        " SELECT 'Natural Gas Ref. CCS' " & s & " ('H11S', 'H71','H72','H73') UNION " &
        " SELECT 'Electricity (HP,Pink)' " & s & " ('H21HP') UNION " &
        " SELECT 'Electricity (HT,CuCl,Pink)' " & s & " ('H21HT', 'H21SI') UNION " &
        " SELECT 'Wind Electrolysis' " & s & " ('H61') UNION " &
        " SELECT 'Wind Elect. Green' " & s & " ('H61G') UNION " &
        " SELECT 'Biomass Gasification' " & s & " ('H80') UNION " &
        " SELECT 'Solar Hydrolysis' " & s & " ('H41','H42') UNION " &
        " SELECT 'Solar Green' " & s & " ('H41G') UNION " &
        " SELECT 'Liquid Fuel Gasif.' " & s & " ('HSD70','HSD85') UNION " &
        " SELECT 'Electricity Mix' " & s & " ('H90','HSD90','H91') UNION" &
        " SELECT 'Electricity Green' " & s & " ('H90G','H91G','HBD90') UNION" &
        " SELECT 'Biomass Gasific. (2nd Gen)' " & s & " ('H80G')"


    End Function

    '19. Hydrogen Production (Detail)
    Function SQLquery19(caseName As String, region As String) As String

        Dim s As String

        s = " AS [Primary Energy], " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
        "   FROM tblTRESULTS " &
        "   WHERE CaseName = '" & caseName & "'" &
        "    AND Parameter IN ('OUTOTH.PRC.LIQ','OUTOTH.PRC.GAS','OUTOTH.PRC.A??','OUTOTH.CON.TOT')" &
        "    AND Region IN (" & region & ") " &
        "    AND Item1 IN "

        SQLquery19 =
        " SELECT  'Coal Gasification'" & s & "('H01')          UNION " &
        " SELECT  'Coal Gas. with CCS' " & s & " ('H01S')      UNION " &
        " SELECT  'Coal Gas. w/CCS Adv.' " & s & " ('H01SF')   UNION " &
        " SELECT  'NGA Steam Ref.' " & s & " ('H11')    UNION " &
        " SELECT  'NGA Steam Ref. with CCS' " & s & " ('H11S')    UNION " &
        " SELECT  'Nucl. HP Electrl.' " & s & " ('H21HP') UNION " &
        " SELECT  'Nucl. HT Electrl.' " & s & " ('H21HT') UNION " &
        " SELECT  'Nucl. CuCl Cycle' " & s & " ('H21SI')       UNION " &
        " SELECT  'Solar Zn/Zn0 to H2' " & s & " ('H41')       UNION SELECT  'Green H2 from Solar' " & s & " ('H41G') UNION " &
        " SELECT 'Solar Coke to H2'  " & s & " ('H42')         UNION " &
        " SELECT 'Central Wind+Electrl.' " & s & " ('H61')     UNION SELECT 'Green Wind+Electrl.' " & s & " ('H61G') UNION " &
        " SELECT 'NGA Autothermal Ref. with CCS' " & s & " ('H71')       UNION " &
        " SELECT 'NGA Chemical Looping Ref. with CCS' " & s & " ('H72')                 UNION SELECT 'NGA Pyrolysis' " & s & " ('H73') UNION " &
        " SELECT 'Biomass Gasification'" & s & " ('H80')       UNION SELECT 'Biomass Gasific. (2nd Gen)'" & s & " ('H80G') UNION " &
        " SELECT 'Alkaline Electrolysis'" & s & "('H90')       UNION " &
        " SELECT 'HP Electrolysis 2020' " & s & " ('H91')      UNION " &
        " SELECT 'Alkaline Elct.(Green)'" & s & "('H90G')      UNION SELECT 'HP Elct. 2020 (Green)' " & s & " ('H91G')    UNION " &
        " SELECT 'Dec. NGA Ref. 1500kg/d' " & s & " ('HBD11')   UNION " &
        " SELECT 'PEM' " & s & " ('HBD90')                     UNION " &
        " SELECT 'Dec. NGA Ref. 100kg/d' " & s & " ('HSD11')   UNION " &
        " SELECT 'Dec. GSL Ref. 470kg/d' " & s & " ('HSD70')   UNION " &
        " SELECT 'Dec. Meth.Ref. 470kg/d' " & s & " ('HSD85')   UNION " &
        " SELECT 'Dec. Electrl. 100kg/d' " & s & " ('HSD90')"

    End Function

    '20. Use of Hydrogen
    Function SQLquery20(caseName As String, region As String) As String

        SQLquery20 =
     "SELECT Technology, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     "FROM " &
     "(SELECT *, " &
     "  IIF(Item1 IN ('I15', 'I23', 'I25'),     'Industry (Heat & Elc)', " &
     "  IIF(Item1 IN ('R12', 'R28'),            'Res&Com. (Heat & Elc)', " &
     "  IIF(Item1 IN ('T17', 'T18'),            'Other Surface Transport',  " &
     "  IIF(Item1 IN ('T2HH', 'T2HF', 'T4HF'),  'Car Transport',            " &
     "  IIF(Item1 IN ('EH2'),                   'Fuel Cell Cogen Grid (Ind. Heat)', " &
      "  IIF(Item1 IN ('EH3'),                  'Fuel Cell Cogen Grid (R&C Heat)',  " &
     "  'Other')))))) AS Technology         " &
     "  FROM tblTRESULTS                  " &
     "  WHERE Parameter = 'FUELUSE.TCH'   " &
     "  AND Item2 = 'HDN'                 " &
     "  AND CaseName = '" & caseName & "' " &
     "  AND Region IN (" & region & ")    " &
     " )GROUP BY Technology"

    End Function

    '21. Electricity Production
    Function SQLquery21(caseName As String, region As String) As String

        Dim s As String
        Dim s2 As String

        s = " AS [Primary Energy], " & template("1/3.6*SUM(T2000) AS 2000", sop, timestep, eop) &
       "   FROM tblTRESULTS " &
       "   WHERE CaseName = '" & caseName & "'" &
       "    AND Parameter IN ('OUTELC.CEN','OUTELC.CPD','OUTELC.DCN') " &
       "    AND Region IN (" & region & ") " &
       "    AND Item1 IN "

        s2 = "  AS [Primary Energy], " & template("1/3.6*SUM(T2000) AS 2000", sop, timestep, eop) &
       "   FROM tblTRESULTS " &
       "   WHERE CaseName = '" & caseName & "'" &
       "    AND Parameter IN ('FUELUSE.TCH') " &
       "    AND Region IN (" & region & ") " &
       "    AND Item2 IN  "


        SQLquery21 =
        " SELECT 'Coal'            " & s & " ('E01','E02','EIG','ES1','H01','E6C') UNION " &
        " SELECT 'Coal (with CCS)' " & s & " ('EC2','EIC','H01SF')           UNION " &
        " SELECT 'Oil'             " & s & " ('E70')                         UNION " &
        " SELECT 'Gas'             " & s & " ('E11','E12','E13','E15','E6A') UNION " &
        " SELECT 'Gas (with CCS)'  " & s & " ('E1C')                         UNION " &
        " SELECT 'Nuclear'         " & s & " ('E21','E22')                   UNION " &
        " SELECT 'Hydro'           " & s & " ('E31')                         UNION " &
        " SELECT 'Solar'           " & s & " ('E41','E41B','E42')                   UNION " &
        " SELECT 'Wind'            " & s & " ('E61','E62')                   UNION " &
        " SELECT 'Biomass'         " & s & " ('BW2', 'BW5', 'BWA','E80','BB1','E82')  UNION " &
        " SELECT 'Biomass (with CCS)' " & s & " ('E83')                      UNION " &
        " SELECT 'Geothermal'      " & s & " ('E81')                         UNION " &
        " SELECT 'Hydrogen CoGen'  " & s & " ('EH2','EH3')                   UNION " &
        " SELECT 'Nuclear (decentral)' " & s2 & " ('ETLH21HP','ETLH21HT', 'ETLH21SI') "



    End Function

    '22. Renewable Electricity Production
    Function SQLquery22(caseName As String, region As String) As String

        Dim s As String

        s = " AS [Primary Energy], " & template("1/3.6*SUM(T2000) AS 2000", sop, timestep, eop) &
       "   FROM tblTRESULTS " &
       "   WHERE CaseName = '" & caseName & "'" &
       "    AND Parameter IN ('OUTELC.CEN','OUTELC.CPD','OUTELC.DCN') " &
       "    AND Region IN (" & region & ") " &
       "    AND Item1 IN "

        SQLquery22 =
        " SELECT 'Hydro'              " & s & " ('E31')       UNION " &
        " SELECT 'Solar'              " & s & " ('E41','E41B','E42') UNION " &
        " SELECT 'Wind'               " & s & " ('E61','E62') UNION " &
        " SELECT 'Biomass'            " & s & " ('BW2', 'BW5', 'BWA','E80','BB1','E82') UNION " &
        " SELECT 'Biomass (with CCS)' " & s & " ('E83')       UNION " &
        " SELECT 'Geothermal'         " & s & " ('E81')"

    End Function

    '23. Electricity CO2 Emissions
    Function SQLquery23(caseName As String, region As String) As String

        Dim s As String, s1 As String, s2 As String, s3 As String, s4 As String, s5 As String, em As String

        s = " Parameter = 'FUELUSE.TCH'         " &
    " AND CaseName = '" & caseName & "' " &
    " AND Region IN (" & region & ")    "

        s1 =
" SELECT 'Coal' AS Tech, " &
    template("(44 / 12) * 0.001 * 0.0258 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('E01','E02','EIG','ES1','H01','EIG','E6C')"
        s2 =
" SELECT 'Coal with CSS' AS Tech, " &
    template("(44 / 12) * 0.001 * 0.0258 * 0.15 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('EC2','EIC','H01SF') "
        s3 =
" SELECT 'Gas' AS Tech, " &
    template("(44 / 12) * 0.001 * 0.0153 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('E11','E12','E13','E15','E6A') "
        s4 =
" SELECT 'Gas with CCS' AS Tech, " &
    template("(44 / 12) * 0.001 * 0.0153 * 0.15 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('E1C') "
        s5 =
" SELECT 'Oil' AS Tech, " &
     template("(44 / 12) * 0.001 * 0.0201 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('E70') "

        em =
" SELECT '1' AS dummy, " & template("SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM (" & s1 & " UNION " & s2 & " UNION " & s3 & " UNION " & s4 & " UNION " & s5 & ")"

        SQLquery23 =
" SELECT 'Intensity' AS Intensity, " & template("t1.T2000/t2.T2000 AS 2000", sop, timestep, eop) &
" FROM " &
" (SELECT '1' AS dummy, " & template("1/3600*SUM(T2000) AS T2000", sop, timestep, eop) &
"  FROM tblTRESULTS " &
"  WHERE CaseName = '" & caseName & "'" &
"    AND Parameter IN ('OUTELC.CEN','OUTELC.CPD','OUTELC.DCN') " &
"    AND Region IN (" & region & ") ) AS t2 " &
" LEFT JOIN " &
" ( " & em & " ) AS t1 ON t1.dummy = t2.dummy"

    End Function

    '24. BEV CO2 Emission
    Function SQLQuery24(caseName As String, region As String) As String

        Dim factors = New Double() {0.7098, 0.7098, 0.6776, 0.6453, 0.6453, 0.6453, 0.6453, 0.6453, 0.6453, 0.6453, 0.6453, 0.6453}
        Dim s As String, s1 As String, s2 As String, s3 As String, s4 As String, s5 As String, em As String
        Dim elcCarEmQuery As String
        Dim carEmQuery As String
        Dim carEmWEURQuery As String
        Dim carEmNAMQuery As String
        Dim carEmASIAQuery As String
        Dim carEmLAFMQuery As String


        s = " Parameter = 'FUELUSE.TCH'         " &
    " AND CaseName = '" & caseName & "' " &
    " AND Region IN (" & region & ")    "

        s1 =
" SELECT 'Coal' AS Tech, " &
    template("(44 / 12) * 0.001 * 0.0258 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('E01','E02','EIG','ES1','H01','EIG','E6C')"
        s2 =
" SELECT 'Coal with CSS' AS Tech, " &
    template("(44 / 12) * 0.001 * 0.0258 * 0.15 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('EC2','EIC','H01SF') "
        s3 =
" SELECT 'Gas' AS Tech, " &
    template("(44 / 12) * 0.001 * 0.0153 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('E11','E12','E13','E15','E6A') "
        s4 =
" SELECT 'Gas with CCS' AS Tech, " &
   template("(44 / 12) * 0.001 * 0.0153 * 0.15 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('E1C') "
        s5 =
" SELECT 'Oil' AS Tech, " &
     template("(44 / 12) * 0.001 * 0.0201 * SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Item1 IN ('E70') "

        em =
" SELECT '1' AS dummy, " & template("SUM(T2000) AS T2000", sop, timestep, eop) &
" FROM (" & s1 & " UNION " & s2 & " UNION " & s3 & " UNION " & s4 & " UNION " & s5 & ")"

        elcCarEmQuery =
" SELECT 'BEV(WtW)' AS Em, " & template("SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM (" &
" SELECT " & template("$*SUM(T2000) AS T2000", sop, timestep, eop, factors) &
" FROM " &
"( SELECT 'Intensity' AS Intensity, " & template("1000/3.6*t1.T2000/t2.T2000 AS T2000", sop, timestep, eop) &
"  FROM " &
"  (SELECT '1' AS dummy, " & template("1/3600*SUM(T2000) AS T2000", sop, timestep, eop) &
"    FROM tblTRESULTS " &
"    WHERE CaseName = '" & caseName & "'" &
"      AND Parameter IN ('OUTELC.CEN','OUTELC.CPD','OUTELC.DCN') " &
"      AND Region IN (" & region & ") ) AS t2 " &
"  LEFT JOIN " &
"  ( " & em & " ) AS t1 ON t1.dummy = t2.dummy " &
") )"

        carEmQuery =
 " SELECT 'ICE(TtW)' AS Em," & template("44/12*SUM(T2000) AS 2000", sop, 10, eop) &
 " FROM ( SELECT                        " &
 " t1.T2000*t2.T2000/9053.0086 AS T2000, " &
 " t1.T2010*t2.T2010/10938.4662 AS T2010, " &
 " t1.T2020*t2.T2020/13814.6596 AS T2020, " &
 " t1.T2030*t2.T2030/16881.4583 AS T2030, " &
 " t1.T2040*t2.T2040/21036.0567 AS T2040, " &
 " t1.T2050*t2.T2050/26635.8312 AS T2050, " &
 " t1.T2060*t2.T2060/30186.2112 AS T2060, " &
 " t1.T2070*t2.T2070/32465.0104 AS T2070, " &
 " t1.T2080*t2.T2080/34078.7579 AS T2080, " &
 " t1.T2090*t2.T2090/35599.0437 AS T2090, " &
 " t1.T2100*t2.T2100/37692.1482 AS T2100, " &
 " t1.T2110*t2.T2110/37692.1482 AS T2110  "
        carEmQuery = carEmQuery &
 " FROM   tblTData AS t1           " &
 "  INNER JOIN tblTRESULTS AS t2   " &
 "      ON  t1.Item1  = t2.Item1   " &
 "      AND t1.Region = t2.Region  " &
 "  WHERE                          " &
 "      t1.ScenarioName = 'VW5_T4_2'       " &
 "  AND t1.Parameter = 'RAT_ACT'           " &
 "  AND t1.Item2 = 'RATCO2CARS'            " &
 "  AND t2.CaseName = '" & caseName & "'   " &
 "  AND t2.Parameter = 'USENRG.TRN_DMD_DM') "

        carEmWEURQuery =
 " SELECT 'ICE(TtW) WEUR' AS Em," & template("44/12*SUM(T2000) AS 2000", sop, 10, eop) &
 " FROM ( SELECT                        " &
 " t1.T2000*t2.T2000/2667 AS T2000, " &
 " t1.T2010*t2.T2010/t3.T2010 AS T2010, " &
 " t1.T2020*t2.T2020/t3.T2020 AS T2020, " &
 " t1.T2030*t2.T2030/t3.T2030 AS T2030, " &
 " t1.T2040*t2.T2040/t3.T2040 AS T2040, " &
 " t1.T2050*t2.T2050/t3.T2050 AS T2050, " &
 " t1.T2060*t2.T2060/t3.T2060 AS T2060, " &
 " t1.T2070*t2.T2070/t3.T2070 AS T2070, " &
 " t1.T2080*t2.T2080/t3.T2080 AS T2080, " &
 " t1.T2090*t2.T2090/t3.T2090 AS T2090, " &
 " t1.T2100*t2.T2100/t3.T2100 AS T2100, " &
 " t1.T2110*t2.T2110/t3.T2110 AS T2110  "
        carEmWEURQuery = carEmWEURQuery &
 " FROM ( tblTData AS t1           " &
 "  INNER JOIN tblTRESULTS AS t2   " &
 "      ON  t1.Item1  = t2.Item1   " &
 "      AND t1.Region = t2.Region) " &
 "  INNER JOIN tblTData AS t3      " &
 "      ON  t1.Region = t3.Region  " &
 "  WHERE                          " &
 "      t1.ScenarioName = 'VW5_T4_2'       " &
 "  AND t1.Parameter = 'RAT_ACT'           " &
 "  AND t1.Item2 = 'RATCO2CARS'            " &
 "  AND t1.Region IN ('WEUR')              " &
 "  AND t2.CaseName = '" & caseName & "'   " &
 "  AND t2.Parameter = 'USENRG.TRN_DMD_DM' " &
 "  AND t3.ScenarioName = 'VW1'            " &
 "  AND t3.Parameter = 'DEMAND'            " &
 "  AND t3.Item2 = 'T2') "

        carEmNAMQuery =
 " SELECT 'ICE(TtW) NAM' AS Em," & template("44/12*SUM(T2000) AS 2000", sop, 10, eop) &
 " FROM ( SELECT                        " &
 " t1.T2000*t2.T2000/4006.3 AS T2000, " &
 " t1.T2010*t2.T2010/t3.T2010 AS T2010, " &
 " t1.T2020*t2.T2020/t3.T2020 AS T2020, " &
 " t1.T2030*t2.T2030/t3.T2030 AS T2030, " &
 " t1.T2040*t2.T2040/t3.T2040 AS T2040, " &
 " t1.T2050*t2.T2050/t3.T2050 AS T2050, " &
 " t1.T2060*t2.T2060/t3.T2060 AS T2060, " &
 " t1.T2070*t2.T2070/t3.T2070 AS T2070, " &
 " t1.T2080*t2.T2080/t3.T2080 AS T2080, " &
 " t1.T2090*t2.T2090/t3.T2090 AS T2090, " &
 " t1.T2100*t2.T2100/t3.T2100 AS T2100, " &
 " t1.T2110*t2.T2110/t3.T2110 AS T2110  "
        carEmNAMQuery = carEmNAMQuery &
 " FROM ( tblTData AS t1           " &
 "  INNER JOIN tblTRESULTS AS t2   " &
 "      ON  t1.Item1  = t2.Item1   " &
 "      AND t1.Region = t2.Region) " &
 "  INNER JOIN tblTData AS t3      " &
 "      ON  t1.Region = t3.Region  " &
 "  WHERE                          " &
 "      t1.ScenarioName = 'VW5_T4_2'       " &
 "  AND t1.Parameter = 'RAT_ACT'           " &
 "  AND t1.Item2 = 'RATCO2CARS'            " &
 "  AND t1.Region IN ('NAM')               " &
 "  AND t2.CaseName = '" & caseName & "'   " &
 "  AND t2.Parameter = 'USENRG.TRN_DMD_DM' " &
 "  AND t3.ScenarioName = 'VW1'            " &
 "  AND t3.Parameter = 'DEMAND'            " &
 "  AND t3.Item2 = 'T2') "


        carEmASIAQuery =
 " SELECT 'ICE(TtW) ASIA' AS Em," & template("44/12*SUM(T2000) AS 2000", sop, 10, eop) &
 " FROM ( SELECT                        " &
 " t1.T2000*t2.T2000/520.1 AS T2000, " &
 " t1.T2010*t2.T2010/t3.T2010 AS T2010, " &
 " t1.T2020*t2.T2020/t3.T2020 AS T2020, " &
 " t1.T2030*t2.T2030/t3.T2030 AS T2030, " &
 " t1.T2040*t2.T2040/t3.T2040 AS T2040, " &
 " t1.T2050*t2.T2050/t3.T2050 AS T2050, " &
 " t1.T2060*t2.T2060/t3.T2060 AS T2060, " &
 " t1.T2070*t2.T2070/t3.T2070 AS T2070, " &
 " t1.T2080*t2.T2080/t3.T2080 AS T2080, " &
 " t1.T2090*t2.T2090/t3.T2090 AS T2090, " &
 " t1.T2100*t2.T2100/t3.T2100 AS T2100, " &
 " t1.T2110*t2.T2110/t3.T2110 AS T2110  "
        carEmASIAQuery = carEmASIAQuery &
 " FROM ( tblTData AS t1           " &
 "  INNER JOIN tblTRESULTS AS t2   " &
 "      ON  t1.Item1  = t2.Item1   " &
 "      AND t1.Region = t2.Region) " &
 "  INNER JOIN tblTData AS t3      " &
 "      ON  t1.Region = t3.Region  " &
 "  WHERE                          " &
 "      t1.ScenarioName = 'VW5_T4_2'       " &
 "  AND t1.Parameter = 'RAT_ACT'           " &
 "  AND t1.Item2 = 'RATCO2CARS'            " &
 "  AND t1.Region IN ('ASIA')              " &
 "  AND t2.CaseName = '" & caseName & "'   " &
 "  AND t2.Parameter = 'USENRG.TRN_DMD_DM' " &
 "  AND t3.ScenarioName = 'VW1'            " &
 "  AND t3.Parameter = 'DEMAND'            " &
 "  AND t3.Item2 = 'T2')   "


        carEmLAFMQuery =
 " SELECT 'ICE(TtW) LAFM' AS Em," & template("44/12*SUM(T2000) AS 2000", sop, 10, eop) &
 " FROM ( SELECT                        " &
 " t1.T2000*t2.T2000/1033.9 AS T2000, " &
 " t1.T2010*t2.T2010/t3.T2010 AS T2010, " &
 " t1.T2020*t2.T2020/t3.T2020 AS T2020, " &
 " t1.T2030*t2.T2030/t3.T2030 AS T2030, " &
 " t1.T2040*t2.T2040/t3.T2040 AS T2040, " &
 " t1.T2050*t2.T2050/t3.T2050 AS T2050, " &
 " t1.T2060*t2.T2060/t3.T2060 AS T2060, " &
 " t1.T2070*t2.T2070/t3.T2070 AS T2070, " &
 " t1.T2080*t2.T2080/t3.T2080 AS T2080, " &
 " t1.T2090*t2.T2090/t3.T2090 AS T2090, " &
 " t1.T2100*t2.T2100/t3.T2100 AS T2100, " &
 " t1.T2110*t2.T2110/t3.T2110 AS T2110  "
        carEmLAFMQuery = carEmLAFMQuery &
 " FROM ( tblTData AS t1           " &
 "  INNER JOIN tblTRESULTS AS t2   " &
 "      ON  t1.Item1  = t2.Item1   " &
 "      AND t1.Region = t2.Region) " &
 "  INNER JOIN tblTData AS t3      " &
 "      ON  t1.Region = t3.Region  " &
 "  WHERE                          " &
 "      t1.ScenarioName = 'VW5_T4_2'       " &
 "  AND t1.Parameter = 'RAT_ACT'           " &
 "  AND t1.Item2 = 'RATCO2CARS'            " &
 "  AND t1.Region IN ('LAFM')              " &
 "  AND t2.CaseName = '" & caseName & "'   " &
 "  AND t2.Parameter = 'USENRG.TRN_DMD_DM' " &
 "  AND t3.ScenarioName = 'VW1'            " &
 "  AND t3.Parameter = 'DEMAND'            " &
 "  AND t3.Item2 = 'T2')   "

        SQLQuery24 = elcCarEmQuery & " UNION " & carEmWEURQuery & " UNION " &
             carEmASIAQuery & " UNION " & carEmNAMQuery & " UNION " &
             carEmLAFMQuery & " UNION " & carEmQuery

    End Function

    '25. Oil Depletion
    Function sqlQuery25(caseName As String, region As String, scenarioName As String) As String

        Dim s0 As String
        Dim s1 As String
        Dim s2 As String
        Dim s3 As String
        Dim s4 As String

        s0 =
"SELECT * FROM ( " &
"SELECT Technology, " & template("0.001 * SUM(T2000) AS 2000", sop, 10, eop) &
"FROM ( "
        s1 =
" SELECT Technology,                                            " &
"  y0 AS T2010,                                                 " &
"  y0 - 10*( 3/4*y1+1/4*y2                           ) AS T2020," &
"  y0 - 10*(                 y1+y2                   ) AS T2030," &
"  y0 - 10*( 3/4*y3+1/4*y4  +y1+y2                   ) AS T2040," &
"  y0 - 10*(                 y1+y2+y3+y4             ) AS T2050," &
"  y0 - 10*( 3/4*y5+1/4*y6  +y1+y2+y3+y4             ) AS T2060," &
"  y0 - 10*(                 y1+y2+y3+y4+y5+y6       ) AS T2070," &
"  y0 - 10*( 3/4*y7+1/4*y8  +y1+y2+y3+y4+y5+y6       ) AS T2080," &
"  y0 - 10*(                 y1+y2+y3+y4+y5+y6+y7+y8 ) AS T2090," &
"  y0 - 10*( 3/4*y9+1/4*y10 +y1+y2+y3+y4+y5+y6+y7+y8 ) AS T2100," &
"  y0 - 10*(                 y1+y2+y3+y4+y5+y6+y7+y8+y9+y10 ) AS T2110 " &
" FROM ( "
        s2 =
" SELECT " &
"  IIF(t1.Item1 = 'MINOI11', 'Category I',   " &
"  IIF(t1.Item1 = 'MINOI21', 'Category II',  " &
"  IIF(t1.Item1 = 'MINOI31', 'Category III', " &
"  IIF(t1.Item1 = 'MINOI41', 'Category IV',  " &
"  IIF(t1.Item1 = 'MINOI51', 'Category V',   " &
"  IIF(t1.Item1 = 'MINOI61', 'Category VI', t1.Item1 )))))) AS Technology, " &
"   t1.Value AS y0,                          " &
IIf(sop < 2020, "  IIF(ISNULL(t2.T2010),0,t2.T2010)", " 0") & " AS y1,  " &
"   IIF(ISNULL(t2.T2020),0,t2.T2020) AS y2,  " &
"   IIF(ISNULL(t2.T2030),0,t2.T2030) AS y3,  " &
"   IIF(ISNULL(t2.T2040),0,t2.T2040) AS y4,  " &
"   IIF(ISNULL(t2.T2050),0,t2.T2050) AS y5,  " &
"   IIF(ISNULL(t2.T2060),0,t2.T2060) AS y6,  " &
"   IIF(ISNULL(t2.T2070),0,t2.T2070) AS y7,  " &
"   IIF(ISNULL(t2.T2080),0,t2.T2080) AS y8,  " &
"   IIF(ISNULL(t2.T2090),0,t2.T2090) AS y9, " &
"   IIF(ISNULL(t2.T2100),0,t2.T2100) AS y10, " &
"   IIF(ISNULL(t2.T2110),0,t2.T2110) AS y11  "

        s3 =
 " FROM tblTIDDATA AS t1                            " &
 " LEFT JOIN                                        " &
 " (SELECT * FROM tblTRESULTS                       " &
 "   WHERE Parameter = 'RESOURCE.L'                 " &
 "     AND CaseName  = '" & caseName & "' ) AS t2   " &
 " ON t1.Item1 = t2.Item1                           " &
 "   AND t1.Region = t2.Region                      " &
 " WHERE t1.Parameter = 'CUM'                       " &
 "  AND t1.ScenarioName IN ('" & scenarioName & "') " &
 "  AND t1.Region IN (" & region & ")               " &
 "  AND t1.Item1 LIKE 'MINOI%' "

        s4 = " ) ) GROUP BY Technology ) WHERE Technology <> 'Category IV' "

        sqlQuery25 = s0 & s1 & s2 & s3 & s4

    End Function

    '26. Oil Production
    Function sqlQuery26(caseName As String) As String

        sqlQuery26 =
" SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM tblTRESULTS                                     " &
" WHERE Parameter = 'OUTOTH.PRC.LIQ'                   " &
"  AND Item1 IN ('S13','S14','S15','S16','S17')        " &
"  AND CaseName  = '" & caseName & "'                  " &
" GROUP BY Region "

    End Function

    '27. Gas Production
    Function sqlQuery27(caseName As String) As String

        sqlQuery27 =
" SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM tblTRESULTS                                     " &
" WHERE Parameter = 'OUTOTH.PRC.GAS'                   " &
"  AND Item1 IN ('S18','S19','S1A','S1B','S1C','S1D')  " &
"  AND CaseName  = '" & caseName & "'                  " &
" GROUP BY Region "

    End Function

    '28. Coal Production
    Function sqlQuery28(caseName As String) As String

        sqlQuery28 =
" SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM tblTRESULTS                                     " &
" WHERE Parameter = 'OUTOTH.PRC.SLD'                   " &
"  AND Item1 IN ('S11','S12')  " &
"  AND CaseName  = '" & caseName & "'                  " &
" GROUP BY Region "

    End Function

    '29. Biomass Production
    Function sqlQuery29(caseName As String, region As String) As String

'sqlQuery29 = _
"SELECT * FROM " & _
"(SELECT Technology, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) & _
" FROM (SELECT *,                          " & _
"   IIF(Item1 = 'BWTD', 'Wood',            " & _
"   IIF(Item1 = 'OCTD', 'Oil Crops',       " & _
"   IIF(Item1 = 'CGTD', 'Corn Grains',     " & _
"   IIF(Item1 = 'SUTD', 'Sugar Plants',    " & _
"   IIF(Item1 = 'STTD', 'Stover',          " & _
"   IIF(Item1 = 'DWTD', 'Waste', 'delete')))))) AS Technology " & _
" FROM tblTRESULTS                         " & _
"  WHERE CaseName = '" & caseName & "'     " & _
"    AND Parameter IN ('OUTOTH.PRC.A??')   " & _
"    AND Region IN (" & region & ")        " & _
" ) GROUP BY Technology) WHERE Technology <> 'delete' "

sqlQuery29 =
" SELECT Region, " & template("0.001*sum(T2000) AS 2000", sop, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE CaseName = '" & caseName & "'" &
"   AND Region <> 'REGION0' " &
"   AND( Parameter IN ('OUTOTH.PRC.A??')" &
"           AND Item1 IN ('BWTD','BL26', 'OCTD','CGTD','SUTD','STTD','DWTD' ) AND Item2 NOT LIKE 'ETL%' )" &
" GROUP BY Region"




    End Function

    '30. Cost of Energy Carriers
    Function sqlQuery30(caseName As String, region As String) As String

        Dim s As String, s1 As String, s2 As String, s3 As String, s4 As String, s5 As String, s6 As String
        Dim s7 As String, s8 As String, s9 As String, s10 As String, s11 As String, s12 As String
        Dim s13 As String, s14 As String, s15 As String, s16 As String, s17 As String
        Dim s18 As String, s19 As String, s20 As String

        s = " CaseName = '" & caseName & "' AND Region IN (" & region & ") "

        s1 =
" SELECT 'Gasoline (LDV, b)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M' AND Item2 = 'GLR'"
        s2 =
" SELECT 'Natural Gas (LDV, b)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M' AND Item2 = 'CGR'"
        s3 =
" SELECT 'Diesel (LDV, b)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M' AND Item2 = 'DLR'"
        s4 =
" SELECT 'Hydrogen' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M' AND Item2 = 'HDN'"
        s5 =
" SELECT 'Electricity (m)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ELC.M'  AND Item2 = 'ELC'"
        s6 =
" SELECT 'Oil (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'MR_GTRD(ENT).M'  AND Item2 = 'OIX'"
        s7 =
" SELECT 'Natural Gas (t,a)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M' AND Item2 IN ('NGG', 'NGI')"
        s8 =
" SELECT 'Coal (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'MR_GTRD(ENT).M'  AND Item2 = 'HCX'"
        s9 =
" SELECT 'Diesel (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'MR_GTRD(ENT).M'  AND Item2 = 'DSX'"
        s10 =
" SELECT 'Bio Gas (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'MR_GTRD(ENT).M'  AND Item2 = 'BNGX'"
        s11 =
" SELECT 'Bio Methanol (a)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M'  AND Item2 = 'BME'"
        s12 =
" SELECT 'Bio Ethanol (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M'  AND Item2 = 'BETX'"
        s13 =
" SELECT 'Biofuel(Diesel) (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M'  AND Item2 = 'BBDX'"
        s14 =
" SELECT 'Methanol (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M'  AND Item2 = 'BMTX'"
        s15 =
" SELECT 'Gasoline (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'MR_GTRD(ENT).M'  AND Item2 = 'GSX'"
        s16 =
" SELECT 'Green Hydrogen' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M' AND Item2 = 'HDG'"
        s17 =
" SELECT 'Green Liquid' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M' AND Item2 = 'FTG'"
        s18 =
" SELECT 'Green Liquid (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'FUEL.ENC.M'  AND Item2 = 'FGX'"
        s19 =
" SELECT 'Green Hydrogen (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'MR_GTRD(ENT).M'  AND Item2 = 'HGX'"
        s20 =
" SELECT 'Hydrogen (t)' AS Fuel, " & template("AVG(T2000) AS 2000", sop + timestep, timestep, eop) &
" FROM tblTRESULTS " &
" WHERE " & s & " AND Parameter = 'MR_GTRD(ENT).M'  AND Item2 = 'HDX'"


        sqlQuery30 = s1 & " UNION " & s2 & " UNION " & s3 & " UNION " &
             s4 & " UNION " & s5 & " UNION " & s6 & " UNION " &
             s7 & " UNION " & s8 & " UNION " & s9 & " UNION " &
             s10 & " UNION " & s11 & " UNION " & s12 & " UNION " &
             s13 & " UNION " & s14 & " UNION " & s15 & " UNION " &
             s16 & " UNION " & s17 & " UNION " & s18 & " UNION " &
             s19 & " UNION " & s20

    End Function


    '31. Detailed Fuels in T1 and T3
    Function sqlQuery31(caseName As String, region As String) As String

        Dim s1 As String
        Dim s2 As String
        Dim s3 As String


        'Splitted because too many cagetories:

        s1 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  SWITCH(" &
     "   Item1 = 'BL12', 'Bio-Syngas',              " &
     "   Item1 = 'BL13', 'Natural Gas',             " &
     "   Item1 IN ('BL14', 'BL18'), 'Bio-Ethanol',  " &
     "   Item1 IN ('BL15', 'BL19'), 'Bio-Methanol', " &
     "   Item1 IN ('BL16', 'BL20'), 'Methanol',     " &
     "   Item1 IN ('BL29'), 'Green Liquid',         " &
     "   Item1 = 'BL17', 'Gasoline',                " &
     "   Item1 = 'BL21', 'Diesel(+Fossil-FT)',      " &
     "   Item1 = 'BL22', 'Biofuel(Diesel)',         " &
     "   Item1 = 'BL25', 'FT-Diesel',               " &
     "   1=1, 'dummy1') AS Fuel                     " &
     " FROM tblTRESULTS                        " &
     " WHERE CaseName = '" & caseName & "'     " &
     "   AND Parameter = 'ACTIVITY.L'          " &
     "   AND Region IN (" & region & ")        " &
     "   AND ( Item1 IN ('BL12', 'BL13', 'BL14', 'BL15', 'BL16', 'BL17', 'BL18', 'BL19', 'BL20', 'BL21', 'BL22') " &
     "      OR Item1 IN ('BL25', 'BL29')) " &
     ") GROUP BY Fuel"


        s3 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  SWITCH(Item1 = 'BL07', 'Bio-FT-Jetfuel'," &
     "   Item1 = 'BL24', 'FT-Jetfuel', Item1 = 'BL08', 'Jetfuel(+Fossil-FT)', Item1 = 'BL09', 'Jetfuel', " &
     "   Item1 = 'BL31', 'Green Jetfuel',             " &
     "   1=1, 'dummy2') AS Fuel                     " &
     " FROM tblTRESULTS                        " &
     " WHERE CaseName = '" & caseName & "'     " &
     "   AND Parameter = 'ACTIVITY.L'          " &
     "   AND Region IN (" & region & ")        " &
     "   AND Item1 IN ('BL07', 'BL24', 'BL08', 'BL09', 'BL31') " &
     ") GROUP BY Fuel"




        s2 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 = 'HDN', 'Hydrogen',     " &
     "  IIF(Item2 = 'ELC', 'Electricity',  " &
     "    'Other (Coal)')) AS Fuel         " &
     " FROM tblTRESULTS                    " &
     " WHERE CaseName = '" & caseName & "' " &
     "   AND Parameter = 'FUELC_ENT_DM'    " &
     "   AND Item3 IN ('T1','T3')          " &
     "   AND Region IN (" & region & ")    " &
     "   AND Item2 IN ('HDN', 'ELC', 'HCT') " &
     ") GROUP BY Fuel"

        sqlQuery31 = s1 & " UNION " & s3 & " UNION " & s2

    End Function


    '32. Technologies in T1 and T3
    Function sqlQuery32(caseName As String, region As String) As String

        sqlQuery32 =
" SELECT Technology, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM (SELECT *, " &
"  IIF(Item1 = 'T11', 'Coal',         " &
"  IIF(Item1 = 'T12', 'Oil',          " &
"  IIF(Item1 = 'T13', 'Gas',          " &
"  IIF(Item1 = 'T14', 'Electricity',  " &
"  IIF(Item1 = 'T16', 'Alcohol FC',   " &
"  IIF(Item1 = 'T17', 'Hydrogen FC',  " &
"  IIF(Item1 = 'T18', 'Hydrogen ICE', " &
"  IIF(Item1 = 'T31', 'Aviation',     " &
"  IIF(Item1 = 'T32', 'Aviation Adv.', Item1 ))))))))) AS Technology " &
" FROM tblTRESULTS " &
" WHERE CaseName  = '" & caseName & "' " &
"   AND Parameter = 'USENRG.TRN_DMD_DM' " &
"   AND Item2 IN ('T1','T3') " &
"   AND Region IN (" & region & ") " &
") GROUP BY Technology"

    End Function


    '33. CO2-Emissions from Cars
    Function sqlQuery33(caseName As String) As String

        sqlQuery33 =
 " SELECT Region, " & template("44/12*0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
 " FROM tblTRESULTS                     " &
 " WHERE CaseName  = '" & caseName & "' " &
 "   AND Parameter = 'EMISSION.L'       " &
 "   AND Item2     = 'CARCO2'           " &
 " GROUP BY Region                      "

    End Function


    '34. Detailed Fuels in Transport
    Function sqlQuery34(caseName As String, region As String) As String

        Dim s1 As String
        Dim s2 As String
        Dim s3 As String

        s1 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item1 IN ('BL01', 'BL17'), 'Gasoline',             " &
     "  IIF(Item1 IN ('BL08', 'BL09'), 'Jetfuels(+Fossil-FT)', " &
     "  IIF(Item1 IN ('BL07'),         'Bio-Jetfuels',         " &
     "  IIF(Item1 IN ('BL24'),         'FT-Jetfuels',       " &
     "  IIF(Item1 IN ('BL31'),         'Green Jetfuels',         " &
     "  IIF(Item1 IN ('BL02', 'BL14', 'BL18'), 'Bio-Ethanol',  " &
     "  IIF(Item1 IN ('BL03', 'BL15', 'BL19'), 'Bio-Methanol', " &
     "  IIF(Item1 IN ('BL04', 'BL16', 'BL20'), 'Methanol',     " &
     "  IIF(Item1 IN ('BL05', 'BL21'), 'Diesel(+Fossil-FT)',   " &
     "  IIF(Item1 IN ('BL06', 'BL22'), 'Biofuel(Diesel)',      " &
     "  IIF(Item1 IN ('BL23', 'BL25'), 'FT-Liquid',            " &
     "  IIF(Item1 IN ('BL29', 'BL30', 'BL32'), 'Green-Liquid', " &
     "  IIF(Item1 IN ('BL10', 'BL12'), 'Bio-Syngas',           " &
     "  IIF(Item1 IN ('BL11', 'BL13'), 'Natural Gas',          " &
     "    'dummy1')))))))))))))) AS Fuel      " &
     " FROM tblTRESULTS                     " &
     " WHERE CaseName = '" & caseName & "'  " &
     "   AND Parameter = 'ACTIVITY.L'       " &
     "   AND Region IN (" & region & ")     " &
     "   AND Item1 LIKE 'BL%'               " &
     ") GROUP BY Fuel "

        ' splitted because too large:

        s1 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     " SWITCH ( " &
     "  Item1 IN ('BL02', 'BL14', 'BL18'), 'Bio-Ethanol',  " &
     "  Item1 IN ('BL03', 'BL15', 'BL19'), 'Bio-Methanol', " &
     "  Item1 IN ('BL04', 'BL16', 'BL20'), 'Methanol',     " &
     "  Item1 IN ('BL05', 'BL21'), 'Diesel(+Fossil-FT)',   " &
     "  Item1 IN ('BL06', 'BL22'), 'Biofuel(Diesel)',      " &
     "  Item1 IN ('BL23', 'BL25'), 'H2-FT-Diesel',         " &
     "  Item1 IN ('BL29', 'BL30', 'BL32'), 'Green-Liquid', " &
     "  Item1 IN ('BL10', 'BL12'), 'Bio-Syngas',           " &
     "  Item1 IN ('BL11', 'BL13'), 'Natural Gas',          " &
     "  1=1,  'dummy1')                       AS Fuel      " &
     " FROM tblTRESULTS                     " &
     " WHERE CaseName = '" & caseName & "'  " &
     "   AND Parameter = 'ACTIVITY.L'       " &
     "   AND Region IN (" & region & ")     " &
     "   AND ( Item1 IN ('BL02', 'BL03', 'BL04', 'BL05', 'BL06', 'BL23', 'BL29', 'BL10', 'BL11', 'BL14', 'BL15') " &
     "      OR Item1 IN ('BL16', 'BL21', 'BL22', 'BL25', 'BL30', 'BL12', 'BL13') ) " &
     ") GROUP BY Fuel "


        ' gasoline and jet:
        s3 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     " SWITCH ( " &
     "  Item1 IN ('BL01', 'BL17'), 'Gasoline',             " &
     "  Item1 IN ('BL08', 'BL09'), 'Jetfuels(+Fossil-FT)', " &
     "  Item1 IN ('BL07'),         'Bio-Jetfuels',         " &
     "  Item1 IN ('BL24'),         'H2-FT-Jetfuels',       " &
     "  Item1 IN ('BL31'),         'SAF-Jetfuels',         " &
     "  1=2,  'dummy2')                       AS Fuel      " &
     " FROM tblTRESULTS                     " &
     " WHERE CaseName = '" & caseName & "'  " &
     "   AND Parameter = 'ACTIVITY.L'       " &
     "   AND Region IN (" & region & ")     " &
     "   AND Item1 IN ('BL01', 'BL17', 'BL08', 'BL09', 'BL07', 'BL24', 'BL31') " &
     ") GROUP BY Fuel "


        s2 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 = 'HDN', 'Hydrogen',      " &
     "  IIF(Item2 = 'ELC', 'Electricity',   " &
     "    'Other (Coal)')) AS Fuel          " &
     " FROM tblTRESULTS                     " &
     " WHERE CaseName = '" & caseName & "'  " &
     "   AND Parameter = 'FUELC_ENT_DM'     " &
     "   AND Item3 IN ('T1','T2','T3','T4') " &
     "   AND Region IN (" & region & ")     " &
     "   AND Item2 IN ('HDN', 'ELC', 'HCT') " &
     ") GROUP BY Fuel "

        sqlQuery34 = s1 & " UNION " & s3 & " UNION " & s2

    End Function

    '35. Technology Mix for all Cars, Diesel split, by cardinality
    Function SQLquery35(caseName As String, region As String) As String
        Dim mileage As Double

        mileage = 1 / 14
        If UCase(region) = "'ASIAPAC'" Then
            mileage = 1 / 15.6
        End If
        If UCase(region) = "'AUSNZL'" Then
            mileage = 1 / 13.5
        End If
        If UCase(region) = "'BRAZIL'" Then
            mileage = 1 / 15.6
        End If
        If UCase(region) = "'CANMEX'" Then
            mileage = 1 / 14.2
        End If
        If UCase(region) = "'CENASIA'" Then
            mileage = 1 / 13.2
        End If
        If UCase(region) = "'CHINAREG'" Then
            mileage = 1 / 15.2
        End If
        If UCase(region) = "'EEUR'" Then
            mileage = 1 / 12.4
        End If
        If UCase(region) = "'EU31'" Then
            mileage = 1 / 14.0#
        End If
        If UCase(region) = "'INDIA'" Then
            mileage = 1 / 14.7
        End If
        If UCase(region) = "'JPKRTW'" Then
            mileage = 1 / 8.5
        End If
        If UCase(region) = "'LAC'" Then
            mileage = 1 / 15.6
        End If
        If UCase(region) = "'MENA'" Then
            mileage = 1 / 18.2
        End If
        If UCase(region) = "'RUSSIA'" Then
            mileage = 1 / 13.2
        End If
        If UCase(region) = "'SSAFRICA'" Then
            mileage = 1 / 12.4
        End If
        If UCase(region) = "'USA'" Then
            mileage = 1 / 18
        End If

        ' default mileage is on top

        SQLquery35 =
" SELECT Technology, " & template("SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM (SELECT *, " &
template("T2000 * " & mileage & " AS T2000", sop, timestep, eop) & ", " &
"  IIF(Item1 IN ('T2GI', 'T2GA', 'T4GI', 'T4GA'), 'Gasoline Type ICEV',     " &
"  IIF(Item1 IN ('T2DI', 'T2DA'), 'Diesel Type ICEV',              " &
"  IIF(Item1 IN ('T2DH'), 'Diesel Type Hybrid',                    " &
"  IIF(Item1 IN ('T2GH', 'T4GH'),         'Gasoline Type Hybrid',  " &
"  IIF(Item1 IN ('T2EH', 'T4EH'),         'Liquid Fuel Plug-in',   " &
"  IIF(Item1 = 'T2CI',                    'Gas Fuel ICEV',         " &
"  IIF(Item1 = 'T2CH',                    'Gas Fuel Hybrid',       " &
"  IIF(Item1 = 'T2HH',                    'Hydrogen Hybrid',       " &
"  IIF(Item1 = 'T2GF',                    'Gasoline Fuel Cell',    " &
"  IIF(Item1 IN ('T2HF','T4HF'),          'Hydrogen Fuel Cell',    " &
"  IIF(Item1 IN ('T2EB','T4EB'),          'Electric Vehicle'  , Item1))))))))))) AS Technology " &
" FROM tblTRESULTS " &
" WHERE CaseName  = '" & caseName & "' " &
"   AND Parameter = 'USENRG.TRN_DMD_DM' " &
"   AND Item2 IN ('T2','T4') " &
"   AND Region IN (" & region & ") " &
") GROUP BY Technology"





    End Function

    Function dataFromSheet1(region As String, num As Integer, yOffset As Integer) As String

        Dim str As String
        Dim c As Range
        Dim xOffset As Integer
        Dim xStart As Integer
        Dim i As Integer

        c = Worksheets("Sheet1").Range("C1:C20000").Find(num, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
        str = " SELECT TOP 1 '" & region & "' AS Region, "
        xOffset = 0
        xStart = 5 'offset of first column

        For i = sop To eop Step timestep
            str = str & c.Offset(5 + yOffset, xStart + xOffset).Value & " AS " & CStr(i)
            If i < eop Then
                str = str & ", "
            Else
                str = str & " FROM tblTRESULTS "
            End If
            xOffset = xOffset + 1
        Next i

        'MsgBox (str)
        dataFromSheet1 = str

    End Function

    '36. GDP market scenario
    Function SQLquery36(caseName As String) As String

        SQLquery36 =
dataFromSheet1("ASIAPAC", 36, 0) & " UNION " &
dataFromSheet1("AUSNZL", 36, 1) & " UNION " &
dataFromSheet1("BRAZIL", 36, 2) & " UNION " &
dataFromSheet1("CANMEX", 36, 3) & " UNION " &
dataFromSheet1("CENASIA", 36, 4) & " UNION " &
dataFromSheet1("CHINAREG", 36, 5) & " UNION " &
dataFromSheet1("EEUR", 36, 6) & " UNION " &
dataFromSheet1("EU31", 36, 7) & " UNION " &
dataFromSheet1("INDIA", 36, 8) & " UNION " &
dataFromSheet1("JPKRTW", 36, 9) & " UNION " &
dataFromSheet1("LAC", 36, 10) & " UNION " &
dataFromSheet1("MENA", 36, 11) & " UNION " &
dataFromSheet1("RUSSIA", 36, 12) & " UNION " &
dataFromSheet1("SSAFRICA", 36, 13) & " UNION " &
dataFromSheet1("USA", 36, 14)

    End Function


    '37. GDP state scenario
    Function SQLquery37(caseName As String) As String

        SQLquery37 =
dataFromSheet1("ASIAPAC", 37, 0) & " UNION " &
dataFromSheet1("AUSNZL", 37, 1) & " UNION " &
dataFromSheet1("BRAZIL", 37, 2) & " UNION " &
dataFromSheet1("CANMEX", 37, 3) & " UNION " &
dataFromSheet1("CENASIA", 37, 4) & " UNION " &
dataFromSheet1("CHINAREG", 37, 5) & " UNION " &
dataFromSheet1("EEUR", 37, 6) & " UNION " &
dataFromSheet1("EU31", 37, 7) & " UNION " &
dataFromSheet1("INDIA", 37, 8) & " UNION " &
dataFromSheet1("JPKRTW", 37, 9) & " UNION " &
dataFromSheet1("LAC", 37, 10) & " UNION " &
dataFromSheet1("MENA", 37, 11) & " UNION " &
dataFromSheet1("RUSSIA", 37, 12) & " UNION " &
dataFromSheet1("SSAFRICA", 37, 13) & " UNION " &
dataFromSheet1("USA", 37, 14)

    End Function


    '38. Population Jazz scenario
    Function SQLquery38(caseName As String) As String

        SQLquery38 =
dataFromSheet1("ASIAPAC", 38, 0) & " UNION " &
dataFromSheet1("AUSNZL", 38, 1) & " UNION " &
dataFromSheet1("BRAZIL", 38, 2) & " UNION " &
dataFromSheet1("CANMEX", 38, 3) & " UNION " &
dataFromSheet1("CENASIA", 38, 4) & " UNION " &
dataFromSheet1("CHINAREG", 38, 5) & " UNION " &
dataFromSheet1("EEUR", 38, 6) & " UNION " &
dataFromSheet1("EU31", 38, 7) & " UNION " &
dataFromSheet1("INDIA", 38, 8) & " UNION " &
dataFromSheet1("JPKRTW", 38, 9) & " UNION " &
dataFromSheet1("LAC", 38, 10) & " UNION " &
dataFromSheet1("MENA", 38, 11) & " UNION " &
dataFromSheet1("RUSSIA", 38, 12) & " UNION " &
dataFromSheet1("SSAFRICA", 38, 13) & " UNION " &
dataFromSheet1("USA", 38, 14)

    End Function


    '39. Population Symphony scenario
    Function SQLquery39(caseName As String) As String

        SQLquery39 =
dataFromSheet1("ASIAPAC", 39, 0) & " UNION " &
dataFromSheet1("AUSNZL", 39, 1) & " UNION " &
dataFromSheet1("BRAZIL", 39, 2) & " UNION " &
dataFromSheet1("CANMEX", 39, 3) & " UNION " &
dataFromSheet1("CENASIA", 39, 4) & " UNION " &
dataFromSheet1("CHINAREG", 39, 5) & " UNION " &
dataFromSheet1("EEUR", 39, 6) & " UNION " &
dataFromSheet1("EU31", 39, 7) & " UNION " &
dataFromSheet1("INDIA", 39, 8) & " UNION " &
dataFromSheet1("JPKRTW", 39, 9) & " UNION " &
dataFromSheet1("LAC", 39, 10) & " UNION " &
dataFromSheet1("MENA", 39, 11) & " UNION " &
dataFromSheet1("RUSSIA", 39, 12) & " UNION " &
dataFromSheet1("SSAFRICA", 39, 13) & " UNION " &
dataFromSheet1("USA", 39, 14)

    End Function


    ' 40. Fuels in Industry Thermal
    Function SQLquery40(caseName As String, region As String) As String

        SQLquery40 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 IN ('DST'),                 'Oil Type',      " &
     "  IIF(Item2 IN ('ALT', 'BIP'),          'Bio, Alc',      " &
     "  IIF(Item2 IN ('FCH', 'HDN'),          'Hydrogen',      " &
     "  IIF(Item2 IN ('ELC'),                 'Electricity',   " &
     "  IIF(Item2 IN ('NGT'),                 'Gas Type',      " &
     "  IIF(Item2 IN ('HCT'),                 'Coal',          " &
     "  IIF(Item2 IN ('LTH'),                 'Heat',          " &
     "  IIF(Item2 IN ('SOT'),                 'Solar',         " &
     "  'Other')))))))) AS Fuel              " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('I1')                  " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "

    End Function


    ' 41. Fuels in Industry Specific
    Function SQLquery41(caseName As String, region As String) As String

        SQLquery41 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 IN ('DST'),                 'Oil Type',    " &
     "  IIF(Item2 IN ('ALT'),                 'Alc',         " &
     "  IIF(Item2 IN ('HDN'),                 'Hydrogen',    " &
     "  IIF(Item2 IN ('ELC'),                 'Electricity', " &
     "  'Other')))) AS Fuel                  " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('I2')                  " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "

    End Function


    ' 42. Fuels in Feedstocks
    Function SQLquery42(caseName As String, region As String) As String

        SQLquery42 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 IN ('DSF'),                 'Oil Type',  " &
     "  IIF(Item2 IN ('ALF'),                 'Alc',       " &
     "  IIF(Item2 IN ('NGF'),                 'Gas Type',  " &
     "  IIF(Item2 IN ('HCO','HCF'),           'Coal',      " &
     "  'Other')))) AS Fuel                  " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('I3')                  " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "

    End Function

    ' 43. Fuels in Other (Residental & Commercial) Sector Thermal (includes non-commercial bio)
    Function SQLquery43(caseName As String, region As String) As String

        SQLquery43 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item3 IN ('R3'),                  'Bio noncomm.',  " &
     "  IIF(Item2 IN ('DST'),                 'Oil Type',      " &
     "  IIF(Item2 IN ('ALT', 'BIP'),          'Bio, Alc',      " &
     "  IIF(Item2 IN ('FCR', 'HDN'),          'Hydrogen',      " &
     "  IIF(Item2 IN ('ELC'),                 'Electricity',   " &
     "  IIF(Item2 IN ('NGT'),                 'Gas Type',      " &
     "  IIF(Item2 IN ('HCT'),                 'Coal',          " &
     "  IIF(Item2 IN ('LTH'),                 'Heat',          " &
     "  IIF(Item2 IN ('SOT'),                 'Solar',         " &
     "  'Other'))))))))) AS Fuel              " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('R2','R3')             " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "

    End Function

    ' 44. Fuels in Other (Residental & Commercial) Sector Specific
    Function SQLquery44(caseName As String, region As String) As String

        SQLquery44 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 IN ('HDN'),                 'Hydrogen',    " &
     "  IIF(Item2 IN ('ELC'),                 'Electricity', " &
     "  'Other')) AS Fuel                  " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('R1')                  " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "

    End Function

    '45. Gas and Coal Electricity Production
    Function SQLquery45(caseName As String, region As String) As String

        Dim s As String

        s = " AS [Primary Energy], " & template("1/3.6*SUM(T2000) AS 2000", sop, timestep, eop) &
       "   FROM tblTRESULTS " &
       "   WHERE CaseName = '" & caseName & "'" &
       "    AND Parameter IN ('OUTELC.CEN','OUTELC.CPD','OUTELC.DCN') " &
       "    AND Region IN (" & region & ") " &
       "    AND Item1 IN "

        SQLquery45 =
        " SELECT 'Coal Conv'       " & s & " ('E01', 'ES1')   UNION " &
        " SELECT 'Coal Adv'        " & s & " ('E02', 'EC2')   UNION " &
        " SELECT 'Coal IGCC'       " & s & " ('EIG', 'EIC')   UNION " &
        " SELECT 'Coal Cogen'      " & s & " ('E6C', 'H01SF') UNION " &
        " SELECT 'Gas CC'          " & s & " ('E11', 'E1C')   UNION " &
        " SELECT 'Gas Turbine'     " & s & " ('E12')          UNION " &
        " SELECT 'Gas Conv'        " & s & " ('E13')          UNION " &
        " SELECT 'Gas Cogen FC'    " & s & " ('E15','E6A')"

    End Function


    '46. Oil Production by Category
    Function sqlQuery46(caseName As String, region As String) As String

        sqlQuery46 =
" SELECT Category, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM " &
" (SELECT *, " &
"   IIF(Item1 IN ('S13'), 'Category I',    " &
"   IIF(Item1 IN ('S14'), 'Category II',   " &
"   IIF(Item1 IN ('S15'), 'Category III',  " &
"   IIF(Item1 IN ('S16'), 'Category IV',   " &
"   IIF(Item1 IN ('S17'), 'Category V+VI', " &
"   'Other'))))) AS Category  " &
"  FROM tblTRESULTS                                     " &
"  WHERE Parameter = 'OUTOTH.PRC.LIQ'                   " &
"   AND Item1 IN ('S13','S14','S15','S16','S17')        " &
"   AND CaseName  = '" & caseName & "'                  " &
"   AND Region IN (" & region & ")                      " &
" )GROUP BY Category "

    End Function

    '47. Gas Production by Category
    Function sqlQuery47(caseName As String, region As String) As String

        sqlQuery47 =
" SELECT Category, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
" FROM " &
" (SELECT *, " &
"   IIF(Item1 IN ('S18'), 'Category I',   " &
"   IIF(Item1 IN ('S19'), 'Category II',  " &
"   IIF(Item1 IN ('S1A'), 'Category III', " &
"   IIF(Item1 IN ('S1B'), 'Category IV',  " &
"   IIF(Item1 IN ('S1C'), 'Category V',   " &
"   IIF(Item1 IN ('S1D'), 'Category VI',  " &
"   'Other')))))) AS Category  " &
"  FROM tblTRESULTS                                     " &
"  WHERE Parameter = 'OUTOTH.PRC.GAS'                   " &
"   AND Item1 IN ('S18','S19','S1A','S1B','S1C','S1D')  " &
"   AND CaseName  = '" & caseName & "'                  " &
"   AND Region IN (" & region & ")                      " &
" )GROUP BY Category "

    End Function


    '48. Net Import/Export by Region
    Function sqlQuery48(caseName As String) As String

        sqlQuery48 =
" SELECT Region, " & template("SUM(IIF(Item1 LIKE 'EXP%' OR Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_TSEP.L', 'R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L', 'R_TSEPELC.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 NOT LIKE 'ETL%' " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function

    '49. Trade Volume by Region
    Function sqlQuery49(caseName As String) As String

        sqlQuery49 =
" SELECT Region, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_TSEP.L','R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L', 'R_TSEPELC.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 NOT LIKE 'ETL%' " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function



    '50. Gas Depletion
    Function sqlQuery50(caseName As String, region As String, scenarioName As String) As String

        Dim s0 As String
        Dim s1 As String
        Dim s2 As String
        Dim s3 As String
        Dim s4 As String

        s0 =
"SELECT * FROM ( " &
"SELECT Technology, " & template("0.001 * SUM(T2000) AS 2000", sop, 10, eop) &
"FROM ( "
        s1 =
" SELECT Technology,                                            " &
"  y0 AS T2010,                                                 " &
"  y0 - 10*( 3/4*y1+1/4*y2                           ) AS T2020," &
"  y0 - 10*(                 y1+y2                   ) AS T2030," &
"  y0 - 10*( 3/4*y3+1/4*y4  +y1+y2                   ) AS T2040," &
"  y0 - 10*(                 y1+y2+y3+y4             ) AS T2050," &
"  y0 - 10*( 3/4*y5+1/4*y6  +y1+y2+y3+y4             ) AS T2060," &
"  y0 - 10*(                 y1+y2+y3+y4+y5+y6       ) AS T2070," &
"  y0 - 10*( 3/4*y7+1/4*y8  +y1+y2+y3+y4+y5+y6       ) AS T2080," &
"  y0 - 10*(                 y1+y2+y3+y4+y5+y6+y7+y8 ) AS T2090," &
"  y0 - 10*( 3/4*y9+1/4*y10 +y1+y2+y3+y4+y5+y6+y7+y8 ) AS T2100," &
"  y0 - 10*(                 y1+y2+y3+y4+y5+y6+y7+y8+y9+y10 ) AS T2110 " &
" FROM ( "
        s2 =
" SELECT " &
"  IIF(t1.Item1 = 'MINNG11', 'Category I',   " &
"  IIF(t1.Item1 = 'MINNG21', 'Category II',  " &
"  IIF(t1.Item1 = 'MINNG31', 'Category III', " &
"  IIF(t1.Item1 = 'MINNG41', 'Category IV',  " &
"  IIF(t1.Item1 = 'MINNG51', 'Category V',   " &
"  IIF(t1.Item1 = 'MINNG61', 'Category VI', t1.Item1 )))))) AS Technology, " &
"   t1.Value AS y0,                          " &
IIf(sop < 2020, "  IIF(ISNULL(t2.T2010),0,t2.T2010)", " 0") & " AS y1,  " &
"  IIF(ISNULL(t2.T2020),0,t2.T2020) AS y2,  " &
"  IIF(ISNULL(t2.T2030),0,t2.T2030) AS y3,  " &
"  IIF(ISNULL(t2.T2040),0,t2.T2040) AS y4,  " &
"  IIF(ISNULL(t2.T2050),0,t2.T2050) AS y5,  " &
"  IIF(ISNULL(t2.T2060),0,t2.T2060) AS y6,  " &
"  IIF(ISNULL(t2.T2070),0,t2.T2070) AS y7,  " &
"  IIF(ISNULL(t2.T2080),0,t2.T2080) AS y8,  " &
"  IIF(ISNULL(t2.T2090),0,t2.T2090) AS y9,  " &
"  IIF(ISNULL(t2.T2100),0,t2.T2100) AS y10, " &
"  IIF(ISNULL(t2.T2110),0,t2.T2110) AS y11  "

        s3 =
 " FROM tblTIDDATA AS t1                            " &
 " LEFT JOIN                                        " &
 " (SELECT * FROM tblTRESULTS                       " &
 "   WHERE Parameter = 'RESOURCE.L'                 " &
 "     AND CaseName  = '" & caseName & "' ) AS t2   " &
 " ON t1.Item1 = t2.Item1                           " &
 "   AND t1.Region = t2.Region                      " &
 " WHERE t1.Parameter = 'CUM'                       " &
 "  AND t1.ScenarioName = '" & scenarioName & "'    " &
 "  AND t1.Region IN (" & region & ")               " &
 "  AND t1.Item1 LIKE 'MINNG%' "

        s4 =
 " )  ) GROUP BY Technology ) WHERE Technology <> 'Category IV' "

        sqlQuery50 = s0 & s1 & s2 & s3 & s4

    End Function


    ' 51. TFC
    Function SQLquery51(caseName As String, region As String) As String


        Dim s0 As String
        Dim s1 As String
        Dim s2 As String
        'Dim s3 As String


        s0 =
     " SELECT *, " &
     "  IIF(Item3 IN ('R3'),                                'Bio noncomm.',  " &
     "  IIF(Item2 IN ('GLR','DLR','DSF','DST','OPR'),       'Oil Type',      " &
     "  IIF(Item2 IN ('ALR','ALT','ALF','BIP'),             'Bio, Alc',      " &
     "  IIF(Item2 IN ('FCR','FCH','HDN'),                   'Hydrogen',      " &
     "  IIF(Item2 IN ('ELC'),                               'Electricity',   " &
     "  IIF(Item2 IN ('NGF','NGT','CGR','CGR1'),            'Gas',           " &
     "  IIF(Item2 IN ('HCO','HCT','HCF'),                   'Coal',          " &
     "  IIF(Item2 IN ('LTH'),                               'Heat',          " &
     "  IIF(Item2 IN ('SOT'),                               'Solar',         " &
     "  'Other (non transport)'))))))))) AS Fuel              " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND Item3 NOT LIKE 'T%'              " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       "


        s1 =
     " SELECT *, " &
     "  IIF(Item1 IN ('BL01', 'BL17'), 'Oil Type',                " &
     "  IIF(Item1 IN ('BL08', 'BL09', 'BL24', 'BL31'), 'Oil Type'," &
     "  IIF(Item1 IN ('BL07'),         'Bio, Alc',                " &
     "  IIF(Item1 IN ('BL02', 'BL14', 'BL18'), 'Bio, Alc',        " &
     "  IIF(Item1 IN ('BL03', 'BL15', 'BL19'), 'Bio, Alc',        " &
     "  IIF(Item1 IN ('BL04', 'BL16', 'BL20'), 'Oil Type',        " &
     "  IIF(Item1 IN ('BL05', 'BL21'), 'Oil Type',                " &
     "  IIF(Item1 IN ('BL06', 'BL22'), 'Bio, Alc',                " &
     "  IIF(Item1 IN ('BL10', 'BL12'), 'Bio, Alc',                " &
     "  IIF(Item1 IN ('BL11', 'BL13'), 'Gas',                     " &
     "    'Other (transport)')))))))))) AS Fuel          " &
     " FROM tblTRESULTS                     " &
     " WHERE CaseName = '" & caseName & "'  " &
     "   AND Parameter = 'ACTIVITY.L'       " &
     "   AND Region IN (" & region & ")     " &
     "   AND Item1 LIKE 'BL%'               "

        s2 =
     " SELECT *, " &
     "  IIF(Item2 = 'HDN', 'Hydrogen',      " &
     "  IIF(Item2 = 'ELC', 'Electricity',   " &
     "    'Coal'))        AS Fuel          " &
     " FROM tblTRESULTS                     " &
     " WHERE CaseName = '" & caseName & "'  " &
     "   AND Parameter = 'FUELC_ENT_DM'     " &
     "   AND Item3 IN ('T1','T2','T3','T4') " &
     "   AND Region IN (" & region & ")     " &
     "   AND Item2 IN ('HDN', 'ELC', 'HCT') "

        SQLquery51 = " SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
             " FROM (" & s0 & " UNION " & s1 & " UNION " & s2 & ")" &
             " GROUP BY Fuel"

    End Function


    '52. Electricity Capacity
    Function SQLquery52(caseName As String, region As String) As String

        Dim s As String
        Dim s1 As String
        Dim s2 As String
        Dim s3 As String
        Dim s4 As String
        Dim s5 As String

        'Gas CHP (E6A) correction 15 region
        s2 =
  template("T2000 * " &
" IIF(Region = 'CANMEX', 1/4.24, IIF(Region = 'CENASIA', 1/2.486, IIF(Region = 'EU31', 1/1.651, " &
" IIF(Region = 'JPKRTW', 1/1.52, IIF(Region = 'RUSSIA',1/1.195, IIF(Region = 'USA',1/2.55," &
" IIF(Region = 'CENTASIA', 1/2.385, IIF(Region = 'EASTASIA', 1/1.214, IIF(Region = 'EUROPE', 1/1.359," &
" IIF(Region = 'NAM', 1/1.2764," &
" 1.0)))))))))) AS TT2000", sop, timestep, eop)

        'Coal CHP (E6C) correction 15 region
        s3 =
  template("T2000 * " &
" IIF(Region IN ('CENASIA', 'CHINAREG', 'EEUR'), 1.0, IIF(Region = 'EU31', 1/1.587, " &
" IIF(Region = 'JPKRTW', 1/2.614, IIF(Region = 'RUSSIA', 1/1.069, IIF(Region IN ('USA','NAM'), 1/2.158," &
" IIF(Region = 'CENTASIA', 1/1.007, IIF(Region = 'EASTASIA', 1/1.436, IIF(Region = 'EUROPE', 1/1.326," &
" 1/2.0)))))))) AS TT2000", sop, timestep, eop)


        s4 = template("SUM(T2000) AS TT2000", sop, timestep, eop)

        s5 = " AS Technology, "

        s =
       " FROM tblTRESULTS " &
       " WHERE CaseName = '" & caseName & "'" &
       "  AND Parameter = 'CAPACITY.L' " &
       "  AND Region IN (" & region & ") " &
       "  AND Item1 IN "


        s1 =
    " SELECT 'Coal'               " & s5 & s4 & s & " ('E01','E02','EIG','ES1','H01') UNION " &
    " SELECT 'Coal'               " & s5 & s3 & s & " ('E6C')                         UNION " &
    " SELECT 'Coal (with CCS)'    " & s5 & s4 & s & " ('EC2','EIC','H01SF')           UNION " &
    " SELECT 'Oil'                " & s5 & s4 & s & " ('E70')                         UNION " &
    " SELECT 'Gas'                " & s5 & s4 & s & " ('E11','E12','E13','E15')       UNION " &
    " SELECT 'Gas'                " & s5 & s2 & s & " ('E6A')                         UNION " &
    " SELECT 'Gas (with CCS)'     " & s5 & s4 & s & " ('E1C')                         UNION " &
    " SELECT 'Nuclear'            " & s5 & s4 & s & " ('E21','E22')                   UNION " &
    " SELECT 'Hydro'              " & s5 & s4 & s & " ('E31')                         UNION " &
    " SELECT 'Solar'              " & s5 & s4 & s & " ('E41','E41B','E42')                   UNION " &
    " SELECT 'Wind'               " & s5 & s4 & s & " ('E61','E62')                   UNION " &
    " SELECT 'Biomass'            " & s5 & s4 & s & " ('BW2', 'BW5', 'BWA','E80','BB1','E82')  UNION " &
    " SELECT 'Biomass (with CCS)' " & s5 & s4 & s & " ('E83')                         UNION " &
    " SELECT 'Geothermal'         " & s5 & s4 & s & " ('E81')                         UNION " &
    " SELECT 'Hydrogen CoGen'     " & s5 & s4 & s & " ('EH2','EH3')"

        'MsgBox " SELECT [Primary Energy], " & template("SUM(T2000) AS 2000", sop, timestep, eop) & _
        '" FROM (" & s1 & _
        '" ) GROUP BY [Primary Energy]"


        SQLquery52 =
" SELECT Technology, " & template("SUM(TT2000) AS 2000", sop, timestep, eop) &
" FROM ( " & s1 & " ) GROUP BY Technology"


    End Function

    '53. Net Coal-Type Import/Export by Region
    Function sqlQuery53(caseName As String) As String

        sqlQuery53 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('HCX', 'BMTX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '54. Net Gas Import/Export by Region
    Function sqlQuery54(caseName As String) As String

        sqlQuery54 =
" SELECT Region, " & template("SUM(IIF(Item1 LIKE 'EXP%' OR Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE ( " &
"  (Item2 = 'LNG' AND Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')) " &
"  OR Parameter = 'R_TSEP.L') " &
"  AND CaseName  = '" & caseName & "' " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function

    '55. Net Oil-Type Import/Export by Region
    Function sqlQuery55(caseName As String) As String

        sqlQuery55 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('OIX','DSX','GSX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '56. Biofuel Import/Export by Region
    Function sqlQuery56(caseName As String) As String

        sqlQuery56 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('BBDX','BETX','BNGX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    ' 57. TFC by sector
    Function SQLquery57(caseName As String, region As String) As String

        SQLquery57 =
     "SELECT Sector, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item3 IN ('R3'),                 'Bio noncomm.',        " &
     "  IIF(Item3 IN ('R2'),                 'Res./comm. thermal',  " &
     "  IIF(Item3 IN ('R1'),                 'Res./comm. specific', " &
     "  IIF(Item3 IN ('I1'),                 'Industry thermal',    " &
     "  IIF(Item3 IN ('I2'),                 'Industry specific',   " &
     "  IIF(Item3 IN ('I3'),                 'Non-energy',          " &
     "  IIF(Item3 IN ('T1'),                 'Other surface trns.', " &
     "  IIF(Item3 IN ('T2','T4'),            'Personal car trns.',  " &
     "  IIF(Item3 IN ('T3'),                 'Aviation',            " &
     "  'Other'))))))))) AS Sector           " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Sector                     "

    End Function


    ' 58. TFC by region
    Function SQLquery58(caseName As String) As String

        SQLquery58 =
     "SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM tblTRESULTS                     " &
     " WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     " GROUP BY Region                       "

    End Function


    ' 59. TFC Industrial by region
    Function SQLquery59(caseName As String) As String

        SQLquery59 =
     "SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM tblTRESULTS                     " &
     " WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('I1','I2','I3')       " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     " GROUP BY Region                       "

    End Function

    ' 60. TFC Res./Comm. by region
    Function SQLquery60(caseName As String) As String

        SQLquery60 =
     "SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM tblTRESULTS                     " &
     " WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('R1','R2','R3')       " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     " GROUP BY Region                       "

    End Function

    ' 61. Electricity in TFC by region
    Function SQLquery61(caseName As String) As String

        SQLquery61 =
     "SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM tblTRESULTS                     " &
     " WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item2 = 'ELC'                   " &
     "  AND CaseName = '" & caseName & "'   " &
     " GROUP BY Region                      "

    End Function

    ' 62. TFC Feedstocks by region
    Function SQLquery62(caseName As String) As String

        SQLquery62 =
     "SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM tblTRESULTS                     " &
     " WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item2 NOT LIKE 'ETL%'           " &
     "  AND Item3 = 'I3'                    " &
     "  AND CaseName = '" & caseName & "'   " &
     " GROUP BY Region                      "

    End Function
    '63. Fuels in Residential/Commercial total
    Function SQLquery63(caseName As String, region As String) As String
        'Dim SQLquery63_a As String
        'Dim SQLquery63_b As String
        SQLquery63 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item3 IN ('R3'),                  'Bio noncomm.',  " &
     "  IIF(Item2 IN ('DST'),                 'Oil Type',      " &
     "  IIF(Item2 IN ('ALT', 'BIP'),          'Bio, Alc',      " &
     "  IIF(Item2 IN ('FCR', 'HDN'),          'Hydrogen',      " &
     "  IIF(Item2 IN ('ELC'),                 'Electricity',   " &
     "  IIF(Item2 IN ('NGT'),                 'Gas Type',      " &
     "  IIF(Item2 IN ('HCT'),                 'Coal',          " &
     "  IIF(Item2 IN ('LTH'),                 'Heat',          " &
     "  IIF(Item2 IN ('SOT'),                 'Solar',         " &
     "  'Other'))))))))) AS Fuel              " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('R1', 'R2','R3')       " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "

    End Function


    ' 64. Fuels in Industry Total
    Function SQLquery64(caseName As String, region As String) As String

        SQLquery64 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 IN ('DST'),                 'Oil Type',      " &
     "  IIF(Item2 IN ('ALT', 'BIP'),          'Bio, Alc',      " &
     "  IIF(Item2 IN ('FCH', 'HDN'),          'Hydrogen',      " &
     "  IIF(Item2 IN ('ELC'),                 'Electricity',   " &
     "  IIF(Item2 IN ('NGT'),                 'Gas Type',      " &
     "  IIF(Item2 IN ('HCT'),                 'Coal',          " &
     "  IIF(Item2 IN ('LTH'),                 'Heat',          " &
     "  IIF(Item2 IN ('SOT'),                 'Solar',         " &
     "  'Other')))))))) AS Fuel              " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('I1','I2')                  " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "



    End Function

    ' 65. Fuels in Industry and Feedstocks Total
    Function SQLquery65(caseName As String, region As String) As String

        SQLquery65 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 IN ('DST','DSF'),                 'Oil Type',      " &
     "  IIF(Item2 IN ('ALT', 'BIP','ALF'),          'Bio, Alc',      " &
     "  IIF(Item2 IN ('FCH', 'HDN'),          'Hydrogen',      " &
     "  IIF(Item2 IN ('ELC'),                 'Electricity',   " &
     "  IIF(Item2 IN ('NGT','NGF'),                 'Gas Type',      " &
     "  IIF(Item2 IN ('HCT','HCO','HCF'),                 'Coal',          " &
     "  IIF(Item2 IN ('LTH'),                 'Heat',          " &
     "  IIF(Item2 IN ('SOT'),                 'Solar',         " &
     "  'Other')))))))) AS Fuel              " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item3 IN ('I1','I2','I3')                  " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "

    End Function
    ' 66. Electricity in Residential/Commercial by region
    Function SQLquery66(caseName As String) As String

        SQLquery66 =
     "SELECT Region, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM tblTRESULTS                     " &
     " WHERE Parameter = 'FUELC_ENT_DM'     " &
     "  AND Item2 = 'ELC'                   " &
     "  AND Item3 IN ('R1','R2','R3')                  " &
     "  AND CaseName = '" & caseName & "'   " &
     " GROUP BY Region                      "

    End Function
    ' 67. Number of cars by region
    Function SQLquery67(caseName As String) As String
        SQLquery67 =
     "SELECT Region, " & template("1 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM tblTRESULTS                     " &
     " WHERE Parameter = 'USENRG.TRN.DM'     " &
     "  AND Item2 in ('T2','T4')            " &
     "  AND CaseName = '" & caseName & "'    " &
     " GROUP BY Region                       "
    End Function

    '68. Total Net imports by region
    Function sqlQuery68(caseName As String) As String

        sqlQuery68 =
" SELECT Region, " & template("SUM(IIF(Item1 LIKE 'EXP%' OR Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_TSEP.L','R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L', 'R_TSEPELC.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 NOT LIKE 'ETL%' " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function
    '69. Fuels in Electricity Generation
    Function SQLquery69(caseName As String, region As String) As String
        'Dim SQLquery69_a As String
        'Dim SQLquery69_b As String
        SQLquery69 =
     "SELECT Fuel, " & template("0.001 * SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 IN ('SOL','WIN','WIO'),                  'Renewables',  " &
     "  IIF(Item2 IN ('HYD'),                 'Hydro',      " &
     "  IIF(Item2 IN ('BIP','BSTT','BWOT','BDWT'),          'Biomass',      " &
     "  IIF(Item2 IN ('HDN'),          'Hydrogen',      " &
     "  IIF(Item2 IN ('URN'),                 'Nuclear',   " &
     "  IIF(Item2 IN ('NGA'),                 'Gas',      " &
     "  IIF(Item2 IN ('DST'),                 'Oil',          " &
     "  IIF(Item2 IN ('HCT'),                 'Coal',          " &
     "  'Other')))))))) AS Fuel              " &
     "  FROM tblTRESULTS                     " &
     "  WHERE Parameter IN ('FUELUSE.TCH','FUELUSEFEQ.TCH')  " &
     "  AND Item1 IN ('E01','E02','EIG','ES1','H01','E6C','EC2','EIC','H01SF','E70','E11','E12','E13','E15','E6A','E1C','E21','E22','E31','E41','E41B','E42','E61','E62','BW2', 'BW5', 'BWA','E80','BB1','E82','E83','E81','EH2','EH3')       " &
     "  AND Item2 NOT LIKE 'ETL%'            " &
     "  AND CaseName = '" & caseName & "'    " &
     "  AND Region IN (" & region & ")       " &
     " ) GROUP BY Fuel                 "


    End Function

    '70. Net LNG Import/Export by Region
    Function sqlQuery70(caseName As String) As String

        sqlQuery70 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE " &
"  Item2 = 'LNG' " &
"  AND Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L') " &
"  AND CaseName  = '" & caseName & "' " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function

    '71. Net Hydrogen Import/Export by Region
    Function sqlQuery71(caseName As String) As String

        sqlQuery71 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('HDX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '72. Net Diesel-Import/Export by Region
    Function sqlQuery72(caseName As String) As String

        sqlQuery72 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('DSX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '73. Net Gasoline-Import/Export by Region
    Function sqlQuery73(caseName As String) As String

        sqlQuery73 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('GSX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '74. Biodiesel Import/Export by Region
    Function sqlQuery74(caseName As String) As String

        sqlQuery74 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('BBDX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '75. Bio-Ethanol Import/Export by Region
    Function sqlQuery75(caseName As String) As String

        sqlQuery75 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('BETX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function

    '76. Biogas Import/Export by Region
    Function sqlQuery76(caseName As String) As String

        sqlQuery76 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('BNGX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function

    '77. Net FT-Hydrogen-Import/Export by Region
    Function sqlQuery77(caseName As String) As String

        sqlQuery77 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,IIF(Parameter = 'R_GTRD(ENT)imp.L',0.001, 0))*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L', 'R_GTRD(ENT)exp.M')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('FHX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '78. Net FT-Green-Import/Export by Region
    Function sqlQuery78(caseName As String) As String

        'get also the marginal, because there may be cases where exp.L and imp.L are all empty; marginal rows are set to zero

        sqlQuery78 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,IIF(Parameter = 'R_GTRD(ENT)imp.L',0.001, 0))*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L', 'R_GTRD(ENT)exp.M')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('FGX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function

    '79. Net Green Hydrogen Import/Export by Region
    Function sqlQuery79(caseName As String) As String

        sqlQuery79 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,0.001)*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('HGX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '80. Hydrogen Import/Export by Region
    Function sqlQuery80(caseName As String) As String

        'get also the marginal, because there may be cases where exp.L and imp.L are all empty; marginal rows are set to zero

        sqlQuery80 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,IIF(Parameter = 'R_GTRD(ENT)imp.L',0.001, 0))*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L', 'R_GTRD(ENT)exp.M')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('HDX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function


    '81. Net Methanol-Import/Export by Region
    Function sqlQuery81(caseName As String) As String

        'get also the marginal, because there may be cases where exp.L and imp.L are all empty; marginal rows are set to zero

        sqlQuery81 =
" SELECT Region, " & template("SUM(IIF(Parameter = 'R_GTRD(ENT)exp.L',-0.001,IIF(Parameter = 'R_GTRD(ENT)imp.L',0.001, 0))*T2000) AS 2000", sop, timestep, eop) &
"  FROM tblTRESULTS                 " &
"  WHERE Parameter IN ('R_GTRD(ENT)imp.L', 'R_GTRD(ENT)exp.L', 'R_GTRD(ENT)exp.M')" &
"  AND CaseName  = '" & caseName & "' " &
"  AND Item2 IN ('BMTX') " &
"  AND Region <> 'REGION0' " &
" GROUP BY Region "

    End Function

    Function fuelQuery(caseName As String, region As String, T2orT4 As String) As String

        Dim s1 As String
        Dim s2 As String

        s1 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item1 = 'BL01', 'Gasoline',        " &
     "  IIF(Item1 = 'BL02', 'Bio-Ethanol',     " &
     "  IIF(Item1 = 'BL03', 'Bio-Methanol',    " &
     "  IIF(Item1 = 'BL04', 'Methanol',        " &
     "  IIF(Item1 IN ('BL32','BL30'), 'Synfuel Green',   " &
     "  IIF(Item1 IN ('BL33','BL23'), 'Synfuel',         " &
     "  IIF(Item1 = 'BL05', 'Diesel',          " &
     "  IIF(Item1 = 'BL06', 'Biofuel(Diesel)', " &
     "  IIF(Item1 = 'BL10', 'Bio-Syngas',      " &
     "  IIF(Item1 = 'BL11', 'Natural Gas',     " &
     "    'dummy1')))))))))) AS Fuel           " &
     " FROM tblTRESULTS                        " &
     " WHERE CaseName = '" & caseName & "'     " &
     "   AND Parameter = 'ACTIVITY.L'          " &
     "   AND Region IN (" & region & ")        " &
     "   AND Item1 IN ('BL01', 'BL02', 'BL03', 'BL04', 'BL32', 'BL33', 'BL23', 'BL30', 'BL05', 'BL06', 'BL10', 'BL11') " &
     ") GROUP BY Fuel"

        s2 =
     "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT *, " &
     "  IIF(Item2 = 'HDN', 'Hydrogen',     " &
     "  IIF(Item2 = 'ELC', 'Electricity',  " &
     "    'dummy2')) AS Fuel               " &
     " FROM tblTRESULTS                    " &
     " WHERE CaseName = '" & caseName & "' " &
     "   AND Parameter = 'FUELC_ENT_DM'    " &
     "   AND Item3 IN (" & T2orT4 & ")     " &
     "   AND Region IN (" & region & ")    " &
     "   AND Item2 IN ('HDN', 'ELC')       " &
     ") GROUP BY Fuel"

        fuelQuery = s1 & " UNION " & s2

    End Function

    Function fuelQuery_BACKUP(caseName As String, region As String, T2orT4 As String) As String

        Dim s1 As String
        Dim s2 As String

        s1 = "SELECT Fuel, " & template("0.001*SUM(T2000) AS 2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT tRegion, " &
     "  IIF(t2Item1 = 'AL1' OR tItem2 = 'MET', 'Methanol',     " &
     "  IIF(t2Item1 = 'AL2' OR tItem2 = 'ETA', 'Ethanol',      " &
     "  IIF(t2Item1 = 'AL3' OR tItem2 = 'BDI', 'Biofuel (Diesel)',   " &
     "  IIF(t2Item1 = 'AL4' OR tItem2 = 'BSG', 'Bio-SNG',      " &
     "  IIF(t2Item1 = 'AL5' OR tItem2 = 'BBT', 'Bio-Methanol', " &
     "  IIF(tItem2 = 'CNG', 'CNG',                " &
     "  IIF(tItem2 = 'ELC', 'Electricity',        " &
     "  IIF(tItem2 = 'HDN', 'Hydrogen',           " &
     "  'Petroleum/Synfuel')))))))) AS Fuel," &
     template(" IIF(tItem2 = 'ALR', IIF( t3T2000 + t4T2000 + t5T2000 + t6T2000 > 0, " &
                            " tT2000*t2T2000/(t3T2000 + t4T2000 + t5T2000 + t6T2000), 0), " &
                    " tT2000 ) AS T2000", sop, timestep, eop) &
     " FROM " &
     "(SELECT t.Region AS tRegion, t2.Item1 AS t2Item1, t.Item2 AS tItem2, " &
      template(" t.T2000 AS tT2000, " &
               " IIF(t2.T2000 IS NULL, 0, t2.T2000) AS t2T2000," &
               " IIF(t3.T2000 IS NULL, 0, t3.T2000) AS t3T2000," &
               " IIF(t4.T2000 IS NULL, 0, t4.T2000) AS t4T2000," &
               " IIF(t5.T2000 IS NULL, 0, t5.T2000) AS t5T2000," &
               " IIF(t6.T2000 IS NULL, 0, t6.T2000) AS t6T2000 ", sop, timestep, eop)
        s2 = " FROM ((((((tblTRESULTS AS t " &
     " LEFT JOIN  (SELECT * FROM tblTRESULTS WHERE Parameter = 'FUELUSE.TCH' AND Item1 IN ('AL1', 'AL2', 'AL3', 'AL4', 'AL5')) AS t2 " &
     "   ON t.Region = t2.Region AND t.CaseName = t2.CaseName AND MID(t.Item2,1,2) = MID(t2.Item1,1,2))" &
     " LEFT JOIN (SELECT * FROM tblTRESULTS WHERE Parameter = 'FUELUSE.TCH' AND Item1 = 'AL1')  AS t3" &
     "   ON t.Region = t3.Region AND t.CaseName = t3.CaseName)" &
     " LEFT JOIN (SELECT * FROM tblTRESULTS WHERE Parameter = 'FUELUSE.TCH' AND Item1 = 'AL2')  AS t4" &
     "   ON t.Region = t4.Region AND t.CaseName = t4.CaseName)" &
     " LEFT JOIN (SELECT * FROM tblTRESULTS WHERE Parameter = 'FUELUSE.TCH' AND Item1 = 'AL3')  AS t5" &
     "   ON t.Region = t5.Region AND t.CaseName = t5.CaseName)" &
     " LEFT JOIN (SELECT * FROM tblTRESULTS WHERE Parameter = 'FUELUSE.TCH' AND Item1 = 'AL4')  AS t6" &
     "   ON t.Region = t6.Region AND t.CaseName = t6.CaseName)" &
     " LEFT JOIN (SELECT * FROM tblTRESULTS WHERE Parameter = 'FUELUSE.TCH' AND Item1 = 'AL5')  AS t7" &
     "   ON t.Region = t7.Region AND t.CaseName = t7.CaseName)" &
     " WHERE t.CaseName = '" & caseName & "' " &
     "   AND t.Parameter = 'FUELC_ENT_DM' AND t.Item3 IN (" & T2orT4 & ") AND t.Item2 NOT LIKE 'ETL%' " &
     "   AND t.Region IN (" & region & ") " &
     ")) GROUP BY Fuel"

        'MsgBox s1
        fuelQuery_BACKUP = s1 & s2

    End Function

    Function template(str As String, b As Integer, timestep As Integer, e As Integer, Optional f As Variant) As String

        Dim s, str2 As String
        Dim y As Integer
        Dim i, imax As Integer 'index and max-index of the additional multiplicative factors f
        Dim factor As String

        s = " "
        y = b
        i = 1
        imax = 0
        If IsMissing(f) Then
            factor = "1.0"
        ElseIf IsArray(f) Then
            i = LBound(f)
            imax = UBound(f)
        Else
            factor = CStr(f)
        End If
        Do While (y <= e)
            If (y > b) Then
                s = s & ", "
            End If
            If (i <= imax) Then
                factor = CStr(f(i))
                i = i + 1
            End If
            str2 = WorksheetFunction.Substitute(str, "2000", CStr(y))
            s = s & WorksheetFunction.Substitute(str2, "$", factor)
            y = y + timestep
        Loop
        template = s & " "

    End Function

    Sub Show_Panel()

        If UserForm1.Visible = False Then
            UserForm1.Show
        End If

    End Sub


    Public Function MyTranspose(vArr As Object) As Object
        Dim lRow As Long
        Dim lCol As Long
        Dim vNewArray() As Object
        ReDim vNewArray(LBound(vArr, 2) To UBound(vArr, 2), LBound(vArr, 1) To UBound(vArr, 1))
        For lRow = LBound(vArr, 1) To UBound(vArr, 1)
            For lCol = LBound(vArr, 2) To UBound(vArr, 2)
                vNewArray(lCol, lRow) = vArr(lRow, lCol)
            Next
        Next
        MyTranspose = vNewArray
    End Function

    Public Sub importData(answerFileName As String, lQuery As String, lTitle As String,
               lUnit As String, lChart As String, lRange As Range, lType As Excel.XlChartType,
               generateChart As Boolean, caseNameLong As String, lTitleShort As String, Optional numFormat As String = "0.0", Optional dH As Integer = 150, Optional dW As Integer = 200)


        Dim nextFreeRow As Excel.Range
        Dim resizedRange As Excel.Range
        Dim myValues As Variant
        'Dim myValues2  As Variant
        Dim nRows As Integer
        Dim nCols As Integer
        Dim i As Integer
        Dim sFieldname As String
        Dim dataRng As Excel.Range

        'Excel 2007: Provider=Microsoft.ACE.OLEDB.12.0;
        ' Needs:
        ' Excel-> VBA Editor -> Tools -> References "Microsoft Active Data Objects 2.1 Library" (perhaps also: "...multidimensional 2.8 Library")

        nextFreeRow = ActiveSheet.Cells(65536, 17).End(XlDirection.xlUp)
        'MsgBox ActiveSheet.Name
        'MsgBox nextFreeRow.Row
        'MsgBox nextFreeRow.Column
        nextFreeRow = nextFreeRow.Offset(RowOffset:=3)
        nextFreeRow.Value = lTitle & " (Units: " & lUnit & ")"
        nextFreeRow = nextFreeRow.Offset(RowOffset:=1)
        'MsgBox nextFreeRow.Row
        'MsgBox nextFreeRow.Column

        If speedImport Then


            'Dim con         As ADODB.Connection 'early binding
            ''Dim con         As Object 'late binding
            'Dim conStr   As String
            'Dim rs          As ADODB.Recordset 'early binding
            ''Dim rs          As Object 'late binding


            ''Set con = CreateObject("ADODB.connection") 'ADO; late binding
            'Set con = New ADODB.Connection 'early binding
            'conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & answerFileName
            'con.Open conStr
            ''Set rs = CreateObject("ADODB.recordset") 'late binding
            'Set rs = New ADODB.Recordset 'early binding
            With UserForm1.rs
                .Source = lQuery
                .ActiveConnection = UserForm1.con
                .CursorType = adOpenForwardOnly
                .Open
                myValues = .GetRows
                nRows = UBound(myValues, 1) + 1
                nCols = UBound(myValues, 2) + 1
                'MsgBox "Rows returned = " & nRows
                'MsgBox "Cols returned = " & nCols
                'MsgBox CStr(myValues)
                'nextFreeRow.CopyFromRecordset .DataSource
                'WorksheetFunction.Transpose(arr)


                For i = 1 To .Fields.Count - 1
                    sFieldname = .Fields(i).Name
                    nextFreeRow.Offset(ColumnOffset:=i).Value = sFieldname
                    nextFreeRow.Offset(ColumnOffset:=i).Font.Bold = True
                    'MsgBox sFieldname
                Next

                resizedRange = nextFreeRow.Resize(nCols + 1, nRows)
                nextFreeRow = nextFreeRow.Offset(RowOffset:=1)
                dataRng = nextFreeRow.Resize(nCols, nRows)
                'MsgBox "Rows in Range = " & dataRng.Rows.Count
                'MsgBox "Cols in Range = " & dataRng.Columns.Count
                'myValues2 = MyTranspose(myValues)
                dataRng.Value = MyTranspose(myValues)

                'MsgBox .Fields(0).Value
            End With
            ''Set thee cursor location.
            'rs.CursorLocation = 3 'adUseClient on early  binding
            'rs.CursorType = 1 'adOpenKeyset on early  binding
            'rs.Open lQuery, con
            ''Redim the table that will contain the filtered data.
            'ReDim myValues(rs.RecordCount)
            'If Not (rs.EOF And rs.BOF) Then
            ' rs.MoveFirst
            ' Dim dbcol As Integer
            ' dbcol = 0
            ' nextFreeRow.Value = rs(dbcol).Value
            UserForm1.rs.Close
            'Set rs = Nothing
            'con.Close
            'Set con = Nothing

        Else

            Dim oQryTable As Object
            oQryTable = ActiveSheet.QueryTables.Add(
   "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & answerFileName & ";",
   nextFreeRow, lQuery)
            '"Select Region, Item1, Item2, Item3, CaseName, Parameter, T2000 from tblTRESULTS")
            'make room for new rows
            oQryTable.RefreshStyle = xlInsertEntireRows
            oQryTable.Refresh False

   resizedRange = oQryTable.ResultRange

            'MsgBox "Delete Query Table"
            oQryTable.Delete

        End If 'speedup or not

        If Len(resizedRange.Rows(resizedRange.Rows.Count).Cells(1, 1).Value) = 0 Then
            'MsgBox "Warning: Name of last data row is empty; the row will not be charted. Chart = " & lTitle
            resizedRange = resizedRange.Resize(resizedRange.Rows.Count - 1)
            'MsgBox resizedRange.Rows.Count
            'MsgBox resizedRange.Columns.Count
        End If
        'MsgBox "Finished Table!"



        Call permutate(resizedRange, lRange)

        If UserForm1.lookupCheckBox.Checked Then
            Call lookupCol(resizedRange, lRange, lTitleShort, caseNameLong)
        End If

        If generateChart Then
            If StrComp(caseNameLong, "") <> 0 Then
                lTitle = lTitle & Chr(10) & caseNameLong
            End If
            Call addChart(resizedRange, lTitle, lUnit, lRange, lType, lChart, dH, dW)
        End If

        'MsgBox "Finished Chart!"
        resizedRange.Offset(1).Resize(resizedRange.Rows.Count - 1).NumberFormat = numFormat

    End Sub


    Public Sub addChart(lRange As Range, lTitle As String, lUnit As String, lcolorRange As Range, lChartType As Excel.XlChartType, lComment As String, dH As Integer, dW As Integer)

        Dim chtChart As Chart
        Dim lSheet As String

        lSheet = ActiveSheet.Name
        chtChart = ActiveSheet.Shapes.addChart(10, 10, dW, dH).Chart
        'Set chtChart = chtChart.Location(WHERE:=xlLocationAsObject, Name:=lSheet)
        With chtChart
            .DisplayBlanksAs = xlZero
            .ChartType = lChartType
            .SetSourceData(Source:=lRange, PlotBy:=xlRows)
            .HasTitle = True
            'please comment the next line out if you are using Excel 2003
            .SetElement(msoElementChartTitleAboveChart)
            .HasTitle = True
            .ChartTitle.Text = lTitle
            .ChartTitle.Font.Bold = True
            .PlotArea.Interior.ColorIndex = xlNone
            .PlotArea.Border.LineStyle = xlNone
            .ChartGroups(1).GapWidth = 20
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).HasTitle = True
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Characters.Text = lUnit
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Font.Size = 11
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.Font.Size = 10
            .Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).TickLabels.NumberFormatLinked = False
            '.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary).TickLabels.numberFormat = "General"
            .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).TickLabels.NumberFormat = "0"
            .Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).TickLabels.Font.Size = 10
            .Legend.Font.Size = 10
            .ChartTitle.Font.Size = 12
            .ChartArea.AutoScaleFont = False
            .Axes(XlAxisType.xlValue).MajorGridlines.Border.ColorIndex = 48
            .Axes(XlAxisType.xlValue).MajorGridlines.Border.Weight = xlHairline
            .Axes(XlAxisType.xlValue).MajorGridlines.Border.LineStyle = xlDot
            If lComment <> vbNullString Then
                .Axes(XlAxisType.xlCategory).HasTitle = True
                .Axes(XlAxisType.xlCategory).AxisTitle.Text = lComment
                .Axes(XlAxisType.xlCategory).AxisTitle.Font.Size = 7
                .Axes(XlAxisType.xlCategory).AxisTitle.HorizontalAlignment = xlLeft
            End If
            'The Parent property is used to set properties of
            'the ChartObject, size etc.
            'ActiveSheet.Range("A65536").End(xlUp).Row
            '.Parent.Top = lRange.Top
            '.Parent.Left = lRange.Columns(lRange.Columns.Count).Left + 9
            '.Parent.Left = lRange.Columns(1).Left
            'MsgBox lRange.Columns(lRange.Columns.Count).Left
            'MsgBox lRange.Columns(lRange.Columns.Count).Left + 9
            'MsgBox lRange.Columns.Count
            'MsgBox lRange.Left

            If .Legend.LegendEntries.Count = 1 Then
                .Legend.Delete()
            ElseIf .Legend.LegendEntries.Count > 15 Then
                .Legend.Font.Size = 8
            End If

        End With

        Call reColor(lcolorRange, chtChart)
        'Worksheets(lSheet).Activate

    End Sub

    'This example sums the data in the first column of query table one. The sum of the first column is displayed below the data range.
    'Set c1 = Sheets("sheet1").QueryTables(1).ResultRange.Columns(1)
    'c1.Name = "Column1"
    'c1.End(xlDown).Offset(2, 0).Formula = "=sum(Column1)"


    Public Sub ArrangeMyCharts(Optional nColumns As Integer = 2,
                    Optional shift As Integer = 0,
                    Optional nRows As Double = 24,
                    Optional nCols As Double = 8,
                    Optional reverse As Boolean = False)

        Dim iChart As Integer
        Dim nCharts As Integer
        Dim dTop As Integer
        Dim dLeft As Integer
        Dim dHeight As Integer
        Dim dWidth As Integer
        'Dim nColumns As Long

        dHeight = nRows * ActiveSheet.Rows(1).RowHeight ' height of all charts
        dWidth = nCols * ActiveSheet.Columns("A").Width ' width of all charts
        dTop = 0      ' top of first row of charts
        dLeft = shift * dWidth ' left of first column of charts
        '    nColumns = 2   ' number of columns of charts
        nCharts = ActiveSheet.ChartObjects.Count

        If reverse Then
            For iChart = 1 To nCharts
                With ActiveSheet.ChartObjects(iChart)
                    .Height = dHeight
                    .Width = dWidth
                    .Top = dTop + Int(((nCharts - iChart + 1) - 1) / nColumns) * dHeight
                    .Left = dLeft + (((nCharts - iChart + 1) - 1) Mod nColumns) * dWidth
                End With
            Next
        Else
            For iChart = 1 To nCharts
                With ActiveSheet.ChartObjects(iChart)
                    .Height = dHeight
                    .Width = dWidth
                    .Top = dTop + Int((iChart - 1) / nColumns) * dHeight
                    .Left = dLeft + ((iChart - 1) Mod nColumns) * dWidth
                End With
            Next
        End If

    End Sub

    Public Sub AutoScale_Off()

        Dim ws As Worksheet, co As ChartObject, i As Integer
        Dim ch As Chart

        For Each ws In ActiveWorkbook.Worksheets
            ' Go through each worksheet in the workbook
            For Each co In ws.ChartObjects
                'In each chart turn the Auto Scale font feature off
                i = i + 1
                co.Chart.ChartArea.AutoScaleFont = False
            Next co
        Next ws
        For Each ch In ActiveWorkbook.Charts
            'Go through each chart in the workbook
            ch.ChartArea.AutoScaleFont = False
            i = i + 1
        Next
        'MsgBox i & " charts have been altered"
        'Application.DisplayAlerts = True
    End Sub

    Public Sub permutate(r As Range, orderRange As Range)

        Dim TempArray As Object
        'Article ID: 169104 - Last Review: October 22, 2000 - Revision: 1.0
        'XL97: Run-Time Error Using Macro to Add Custom List
        ' -> workaround with a temp array

        If orderRange.Rows.Count > 2 Then
            TempArray = orderRange
            Application.AddCustomList(ListArray:=TempArray)
            r.Sort(Key1:=r.Cells(2, 1), Order1:=xlDescending, Header:=xlYes,
      OrderCustom:=Application.CustomListCount + 1)
            Application.DeleteCustomList Application.CustomListCount
  End If

    End Sub

    Public Sub lookupCol(r As Range, lRange As Range, lTitleShort As String, caseNameLong As String)

        'caseNameLong is usually "", better to use caseName in Import

        Dim tagName
        Dim resizedRng As Excel.Range
        Dim regional As Boolean

        If r.Rows.Count - 1 > 0 Then
            'exclude first (header) row
            resizedRng = r.Offset(1, 0).Resize(r.Rows.Count - 1)

            If (StrComp(CStr(r.Cells(1, 1)), "Region") = 0) Then
                regional = True
                lTitleShort = Right(lTitleShort, Len(lTitleShort) - Len("GLOBAL"))
            Else
                regional = False
            End If

            For Each tagName In resizedRng.Columns(1).Cells
                If regional Then
                    tagName.Offset(0, -2).Value = tagName.Value & lTitleShort
                Else
                    tagName.Offset(0, -2).Value = lTitleShort & "_" & tagName.Value
                End If
                tagName.Offset(0, -1).Value = "'"
            Next tagName

        End If

    End Sub


    Public Sub reColor(ref As Range, cht As Chart)

        Dim ser As Series
        Dim refRow As Range
        Dim val2
        'Dim colArray

        For Each ser In cht.SeriesCollection
            With ser
                For Each refRow In ref
                    If (StrComp(CStr(refRow.Cells(1, 1)), .Name) = 0) Then
                        val2 = refRow.Cells(1, 3).Value
                        If cht.ChartType = XlChartType.xlColumnStacked Then
                            If UserForm1.RGBCheckBox.Checked And Not IsEmpty(val2) Then
                                'colArray = getRGB(CStr(val2))
                                'ser.Interior.Color = RGB(colArray(1), colArray(2), colArray(3))
                                .Interior.ColorIndex = CInt(val2)
                                If refRow.Cells(1, 4).Value = "s" Then
                                    .Fill.Patterned(msoPatternDarkUpwardDiagonal)
                                End If
                            Else
                                .Interior.ColorIndex = CInt(refRow.Cells(1, 2))
                            End If
                        End If
                        If cht.ChartType = XlChartType.xlLineMarkers Then
                            If UserForm1.RGBCheckBox.Checked And Not IsEmpty(val2) Then
                                'colArray = getRGB(CStr(val2))
                                'ser.Border.Color = RGB(colArray(1), colArray(2), colArray(3))
                                'ser.MarkerBackgroundColor = RGB(colArray(1), colArray(2), colArray(3))
                                'ser.MarkerForegroundColor = RGB(colArray(1), colArray(2), colArray(3))
                                .Border.ColorIndex = CInt(val2)
                                .MarkerBackgroundColorIndex = CInt(val2)
                                .MarkerForegroundColorIndex = CInt(val2)
                            Else
                                .Border.ColorIndex = CInt(refRow.Cells(1, 2))
                                .MarkerBackgroundColorIndex = CInt(refRow.Cells(1, 2))
                                .MarkerForegroundColorIndex = CInt(refRow.Cells(1, 2))
                            End If
                        End If
                    End If
                Next refRow
            End With
        Next ser

    End Sub

End Module







