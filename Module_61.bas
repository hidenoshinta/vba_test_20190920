Attribute VB_Name = "Module1"
'Attribute VB_Name = "ALG_Module6_1"
'***************************************************************************************************
' 2018.10.4 VBcode
' 2017.9.3-     by shintaku
'   for Algebraic Spiral Scroll
'  Referance) Trans. of the JAR Vol.11,No.3(1994)pp.337-347, Hirokatsu KOHSOKABE
'   代数螺線を基本にしたスクロール流体機械のスクロール形状に関する研究
'  Study on Scroll Profile based on Algebraic Spiral for Scroll Fluid Machine
'
'   Author:     by shintaku
'***************************************************************************************************

Option Explicit

Dim pi      As Double
Dim dw_n_PI()  As Double      ' 回転角Theta θ=0,π,2π,3π,4π,･･･ 半回転毎のIndex No.を dw_n_PI(j)に格納

Dim dw_deg  As Double           ' [deg] Division width for Wrap angle（>0）
Dim dw      As Double           ' Division width for Wrap angle（>0）
Dim dw_t_deg   As Double        ' Division number for Wrap thickness angle [deg] （>0）
Dim dw_t  As Double             ' Division number for Wrap thickness angle [rad] （>0）

Dim dw_c    As Long             ' Matrix number of colume（>0）
Dim dw_n    As Long             ' Matrix Division number（>0）
Dim dw_n_end As Long            ' Matrix number of wrap end（>0）
Dim div_n   As Long             ' Division number for Chamber Area（>0）
Dim Index_I As Long

Dim Matrix_A() As Double         ' equation Matorix A
Dim Matrix_C() As Double         ' equation Matorix C
Dim Matrix_X() As Variant ' Double         ' equation Matorix X

' 冷媒物性
Dim kappa       As Double       ' Adiabatic coefficient
Dim P_suction   As Double       ' [MPa(abs)] Suction Pressure
Dim P_discharge As Double       ' [MPa(abs)] Discharge Pressure
Dim P_mid       As Double:
Dim P_groove    As Double:    ' pressure :OS backsurface, oil-groove


Dim N_rps As Double             ' [rps] Shaft rotational speed


Dim a       As Double           ' Algebraic constat
Dim k       As Double           ' Algebraic constat
Dim g1      As Double           ' Algebraic constat  xi guzai
Dim g2      As Double           ' Algebraic constat
Dim qq      As Double           ' Algebraic constat
Dim Ro      As Double           ' Obit radius
Dim Hw      As Double           ' Wrap Hight

Dim FS_in_srt_deg   As Double
Dim FS_in_end_deg   As Double
Dim FS_out_srt_deg  As Double
Dim FS_out_end_deg  As Double

Dim OS_in_srt_deg   As Double
Dim OS_in_end_deg   As Double
Dim OS_out_srt_deg  As Double
Dim OS_out_end_deg   As Double

Dim FS_in_srt       As Double
Dim FS_in_end       As Double
Dim FS_out_srt      As Double
Dim FS_out_end      As Double
Dim FS_offset_angle As Double
Dim dx              As Double
Dim dy              As Double

Dim OS_in_srt       As Double:    Dim OS_in_srt_0   As Double
Dim OS_in_end       As Double
Dim OS_out_srt      As Double:    Dim OS_out_srt_0  As Double
Dim OS_out_end      As Double
Dim OS_offset_angle As Double


Dim OS_dia As Double
Dim OS_seal As Double

Dim N_wrap_a() As Long     ' contact point Number of chamber A at the(i)
Dim N_wrap_b() As Long     ' contact point Number of chamber B at the(i)
Dim N_wrap_max As Long

'---------------
' arc of Sucrion inlet part
'---------------
 Dim r_Rfi_c As Double:   Dim angle_Rfi_c As Double   ' Suction Inlet FSin-arc Rarius
   Dim x_Rfi_c As Double:   Dim y_Rfi_c As Double  ' Suction Inlet FSin-arc center
 Dim r_Rfo_c As Double:   Dim angle_Rfo_c As Double   ' Suction Inlet FSin-arc Radius
   Dim x_Rfo_c As Double:   Dim y_Rfo_c As Double  ' Suction Inlet FSout-arc center

'---------------
' Wrap head arc dimension
'---------------
    Dim R_head_r1(9) As Double:
    Dim R_head_xc(9) As Double:         Dim R_head_yc(9)  As Double
    Dim R_head_x1(9) As Double:         Dim R_head_y1(9)  As Double
    Dim R_head_x2(9) As Double:         Dim R_head_y2(9)  As Double

'---------------
' Oil Groove dimension
'---------------
 '-- right : start side / arc1 of Oil Groove
 Dim r1_oilgroove As Double:
 Dim t1_oilgroove As Double       ' t1: oilgroove width
   Dim x1_oilgroove_c As Double:        Dim y1_oilgroove_c As Double

   Dim angle1_oilgroove_0 As Double
   Dim angle1_oilgroove_1 As Double
   Dim angle1_oilgroove_2 As Double

 '-- left  : second start side / arc2 of Oil Groove
 Dim r2_oilgroove As Double:
 Dim t2_oilgroove As Double
   Dim x2_oilgroove_c As Double:        Dim y2_oilgroove_c As Double

   Dim angle2_oilgroove_0 As Double
   Dim angle2_oilgroove_1 As Double
   Dim angle2_oilgroove_2 As Double

 '-- cross points of arc1 and arc2 of Oil Groove
 Dim L_tmp0 As Double:
   Dim angle_tmp0 As Double:            Dim angle_tmp00 As Double
   Dim angle_tmp1a As Double:           Dim angle_tmp2a As Double
   Dim angle_tmp1b As Double:           Dim angle_tmp2b As Double

   Dim tmp_1 As Double:       Dim tmp_q As Double:       Dim wrap_n As Double
   Dim x_tmp(99) As Double:   Dim y_tmp(99) As Double
   Dim r_tmp(99) As Double:   Dim q_tmp(99) As Double

  'Reference point about wrap contact point
   Dim x_rfp(40) As Double:   Dim y_rfp(40) As Double:  Dim q_rfp(40) As Double

  '  oilgroove arc1
   Dim r1_tmp As Double:      Dim q1_tmp As Double:
   Dim x1c_tmp As Double:       Dim y1c_tmp As Double:
   Dim x1_tmp As Double:       Dim y1_tmp As Double:

    Dim angle1_tg_0 As Double:
    Dim angle1_tg_1 As Double:
    Dim angle1_tg_2 As Double:

  '  oilgroove arc2
    Dim angle2_tg_0 As Double:
    Dim angle2_tg_1 As Double:
    Dim angle2_tg_2 As Double:

    Dim angle1_OS_Plate_0 As Double:
    Dim angle1_OS_Plate_1 As Double:
    Dim angle1_OS_Plate_2 As Double


'Dim Alg_Const(2, 50) As Variant

Dim DataSheetName   As String
Dim Data_Strage()   As Variant

Dim DataSheetName_2   As String
'Dim Data_Strage_2()   As Variant

Dim DataSheetName_3   As String
Dim Data_Strage_3()   As Variant

Dim Flag_n1         As Long         '【Cell: Phi_1 読み書き用Flag 】 0="OFF" , 1="ON"
Dim kk_Flag         As Long

' --- Wrap angle φ Region  (Start < end : φ1=Phi_1 < φ2=Phi_2 )
'    Dim the_tmp As Double
    Dim Phi_1 As Double:           Dim Phi_2 As Double:
    Dim Phi_0 As Double:           Dim Phi_00 As Double:

    Dim Phi_1_Amax      As Double:      Dim Phi_1_Amin      As Double
    Dim Phi_2_Amax      As Double:      Dim Phi_2_Amin      As Double
    Dim Phi_1_Bmax      As Double:      Dim Phi_1_Bmin      As Double
    Dim Phi_2_Bmax      As Double:      Dim Phi_2_Bmin      As Double
      Dim Phi_2_Amax_deg  As Double:      Dim Phi_2_Amin_deg  As Double
      Dim Phi_2_Bmax_deg  As Double:      Dim Phi_2_Bmin_deg  As Double

' --- Wrap φ strat & end angle  (CI>C2>C3>C4  φ1>φ2>φ3,φ4 )
    Dim P2_C1 As Double:          Dim P2_C2 As Double
    Dim P2_C3 As Double:          Dim P2_C4 As Double
      Dim P2_C1_deg As Double:      Dim P2_C2_deg As Double
      Dim P2_C3_deg As Double:      Dim P2_C4_deg As Double

' --- Rotation angle Theta θ (CI>C2>C3>C4  the1>the2>the3,the4 )
    Dim The_C1 As Double:         Dim The_C2 As Double      '20171031
    Dim The_C3 As Double:         Dim The_C4 As Double      '20171031
      Dim The_Max As Double:        Dim The_Min As Double     '20171031
      Dim Theta_1 As Double:        Dim Theta_2 As Double     '20171031
      Dim turn_wrap_n As Double

    Dim The_Cx As Double
    Dim dw_n_Cx As Long
    Dim txt_Cx As Variant

' --- Index Number of Wrap strat & end angle
    Dim dw_n_C1 As Long:      Dim dw_n_C2 As Long
    Dim dw_n_C3 As Long:      Dim dw_n_C4 As Long


' --- Wrap angle Theta Region　　1,2巻内側                      '171101
    Dim The_C1n() As Double:     Dim The_C2n() As Double      '171101
    Dim The_C3n() As Double:     Dim The_C4n() As Double      '171101
      Dim dw_n_C1n() As Long:      Dim dw_n_C2n() As Long
      Dim dw_n_C3n() As Long:      Dim dw_n_C4n() As Long


' --- Calculate : Volume A,B at (Phi_1,Phi_2)
'   Rotation and Spiral angels
'
    Dim the() As Double:        ' Dim the_deg() As Double   ' θ rotation angle             20171030
    Dim the_c() As Double:
    Dim x_e() As Double:   Dim y_e() As Double

    Dim phi1() As Double:         Dim phi2() As Double      ' φ1，φ2 Wall spiral angle　  20171030
'    Dim phi2m1() As Double                            ' φ2m1，φ2 counter Wall spiral angle　  20180319

    Dim Phi_c_fi() As Double:      'Dim Phi_c_fo() As Double
    'Dim Phi_c_mi() As Double:      Dim Phi_c_mo() As Double

    Dim jj_a As Long
    Dim jj_b As Long


  ' B chamber
    Dim phi1_v() As Double:
    Dim psi_mi1_v() As Double:      Dim psi_fo1_v() As Double
      Dim xmi() As Double:            Dim ymi() As Double:
      Dim RR_mi() As Double:          'Dim del_Smi() As Double
      Dim xfo() As Double:            Dim yfo() As Double:
      Dim RR_fo() As Double:          'Dim del_Sfo() As Double

' for debug of drowing line

      Dim xfi1() As Double:           Dim yfi1() As Double:   ' For whole Wrap from the head start to the tail end
      Dim xfo1() As Double:           Dim yfo1() As Double:
      Dim xmi1() As Double:           Dim ymi1() As Double:
      Dim xmo1() As Double:           Dim ymo1() As Double:

      Dim xfi2() As Double:           Dim yfi2() As Double:   ' for debug of drowing line
      Dim xfo2() As Double:           Dim yfo2() As Double:
      Dim xmi2() As Double:           Dim ymi2() As Double:
      Dim xmo2() As Double:           Dim ymo2() As Double:

      Dim xfi3() As Double:           Dim yfi3() As Double:   ' for debug of drowing line
      Dim xfo3() As Double:           Dim yfo3() As Double:
      Dim xmi3() As Double:           Dim ymi3() As Double:
      Dim xmo3() As Double:           Dim ymo3() As Double:


'  描画 check用　20180706
      Dim curve_xw() As Double:       Dim curve_yw() As Double:
      Dim x_out() As Double:          Dim y_out() As Double:
      Dim x_in() As Double:           Dim y_in() As Double:


  ' A chamber
    Dim psi_mo1_v() As Double:      Dim psi_fi1_v() As Double
      Dim xfi() As Double:            Dim yfi() As Double:
      Dim RR_fi() As Double:          'Dim del_Sfi() As Double
      Dim xmo() As Double:            Dim ymo() As Double:
      Dim RR_mo() As Double:          'Dim del_Smo() As Double

   ' --- Volume check用
      Dim beta_mi1_v() As Double:
      Dim del_mi1_v() As Double:      Dim del_fo1_v() As Double
      Dim beta_fi1_v() As Double:
      Dim del_fi1_v() As Double:      Dim del_mo1_v() As Double

   ' --- Volume
      Dim V_a() As Double:            Dim V_b() As Double
      Dim V_a_Max As Double:          Dim V_a_Min As Double
      Dim V_b_Max As Double:          Dim V_b_Min As Double

' --- 図心 : center of Chamber fiqure
    Dim xg_a_tmp As Double:           Dim yg_a_tmp As Double         ' A Chamber crescent area 図心
    Dim xg_b_tmp As Double:           Dim yg_b_tmp As Double         ' A Chamber crescent area 図心
    Dim Area_tmp As Double

    Dim xg_a_tmp0 As Double:          Dim yg_a_tmp0 As Double         ' A Chamber crescent area 図心
    Dim xg_b_tmp0 As Double:          Dim yg_b_tmp0 As Double         ' A Chamber crescent area 図心
    Dim Area_tmp0 As Double

    Dim xg_a() As Double: Dim yg_a() As Double:  Dim Sg_a() As Double  ' gravity center of chamverA
    Dim xg_b() As Double: Dim yg_b() As Double:  Dim Sg_b() As Double  ' gravity center of chamverB

    Dim xg_f() As Double: Dim yg_f() As Double:  Dim Sg_f() As Double   ' gravity center of FS Wrap parts
    Dim xg_m() As Double: Dim yg_m() As Double:  Dim Sg_m() As Double   ' gravity center of OS Wrap parts

    Dim xg2_a() As Double: Dim yg2_a() As Double: Dim Sg2_a() As Double    ' FS Chamber Center    20180323
    Dim xg2_b() As Double: Dim yg2_b() As Double: Dim Sg2_b() As Double    ' OS Chamber Center    20180323

'   Dim x_zforce() As Double:      Dim y_zforce() As Double     ' FS Chamber Center    20180323
    Dim x_zforce_a() As Double:      Dim y_zforce_a() As Double     ' FS Chamber Center    20180323
    Dim x_zforce_b() As Double:      Dim y_zforce_b() As Double     ' FS Chamber Center    20180323

   Dim r_zforce_a() As Double:      Dim t_zforce_a() As Double     ' OS r-t cordinate 20180658
   Dim r_zforce_b() As Double:      Dim t_zforce_b() As Double

' --- A,B室内圧力
'
    Dim Press_A() As Double:    Dim Press_B() As Double

' 作用力計算用
' A,B室 接点、動径、
   Dim xmi_c() As Double:     Dim ymi_c() As Double:  Dim RR_mi_c() As Double:
   Dim xfo_c() As Double:     Dim yfo_c() As Double:  Dim RR_fo_c() As Double:
   Dim xfi_c() As Double:     Dim yfi_c() As Double:  Dim RR_fi_c() As Double:
   Dim xmo_c() As Double:     Dim ymo_c() As Double:  Dim RR_mo_c() As Double:

' 作用力計算用　β: bata 包絡線の法線と基本代数螺旋の偏角のなす角度
   Dim beta_fi_c() As Double         ' β1，β2 base spiral angle
   Dim beta_mi_c() As Double         ' β1，β2 base spiral angle

' 作用力計算用  δ: Delta 基本代数螺旋の偏角φと包絡線の偏角ψとの差 δ
   Dim del_fi_c() As Double       ' δfi1，δfi2 angle
   Dim del_mo_c() As Double       ' δmo1，δmo2 angle
   Dim del_mi_c() As Double       ' δmi1，δmi2 angle
   Dim del_fo_c() As Double       ' δfo1，δfo2 angle

' 作用力計算用  : 包絡線の偏角 ψ
   Dim psi_fi_c() As Double       ' ψfi1、ψfi2 wall Contact angle
   Dim psi_mo_c() As Double       ' ψmo1、ψmo2 wall Contact angle
   Dim psi_mi_c() As Double       ' ψmi1、ψmi2 wall Contact angle
   Dim psi_fo_c() As Double       ' ψfo1、ψfo2 wall Contact angle


    Dim alpha_dxy    As Double    ' Angle of OS wrap center offset
    Dim RR_dxy      As Double     ' Length of OS wrap center offset
    Dim LLt_dxy()   As Double     ' Tangensial length of OS wrap center offset
    Dim LLr_dxy()   As Double     ' Radial length of OS wrap center offset

' --- Gas力 -----------------------------------------------------------------
'　t接線方向Gas力  Ft   , 作用範囲 接線方向  Lt
    Dim Lt_A() As Double:       Dim Lt_B() As Double:    Dim Lt_D() As Double
     Dim LLt_A() As Double:       Dim LLt_B() As Double

    Dim Ft_a() As Double:       Dim Ft_b() As Double:    Dim Ft_d() As Double
    Dim Ft_AB() As Double

'　r半径方向Gas力　Fr　　, 作用範囲 半径方向  Lr
    Dim Lr_A() As Double:       Dim Lr_B() As Double:    Dim Lr_D() As Double
     Dim LLr_A() As Double:       Dim LLr_B() As Double

    Dim Fr_A() As Double:       Dim Fr_B() As Double:    Dim Fr_D() As Double
    Dim Fr_AB() As Double

'　z軸方向Gas力　Fz    -20180605
    Dim Fz_A() As Double:  Dim Fz_B() As Double:   '
    Dim Fz_f() As Double:  Dim Fz_m() As Double:
    Dim Fz_Za() As Double: Dim Fz_Zb() As Double   ' Wrap面と背面側の、軸方向ガス力の合力
    Dim Fz_sp() As Double                          ' thrust面の反力、軸方向ガス力の合力

' --- Gas圧力 -----------------------------------------------------------------
'　　軸方向Gas圧力　Pz    -20180605
    Dim Pz_A() As Double:     Dim Pz_B() As Double:
    Dim Pz_f() As Double:     Dim Pz_m() As Double:
    Dim Pz_Za() As Double:    Dim Pz_Zb() As Double:
    Dim DP_tm() As Double:    Dim DP_tf() As Double:   'Differece pressure 20181711

' --- Gas力 Moment-----------------------------------------------------------20171215
'　　t接線方向Gas力Moment  Mmt : Moment arm Lmt
    Dim Lmt_A() As Double:       Dim Lmt_B() As Double:       Dim Lmt_D() As Double
      Dim Lt_AB() As Double:
    Dim Mmt_A() As Double:       Dim Mmt_B() As Double:       Dim Mmt_D() As Double
    Dim Mmt_AB() As Double

'　　r半径方向Gas力Moment  Mmr : Moment arm Lmr
    Dim Lmr_A() As Double:       Dim Lmr_B() As Double:       Dim Lmr_D() As Double
      Dim Lr_AB() As Double:
    Dim Mmr_A() As Double:       Dim Mmr_B() As Double:       Dim Mmr_D() As Double
    Dim Mmr_AB() As Double

'　　z軸方向Gas力Moment  -20180605
   Dim Mzx_f() As Double:   Dim Mzy_f() As Double:
   Dim Mzx_m() As Double:   Dim Mzy_m() As Double:
   Dim Mzx_A() As Double:       Dim Mzy_A() As Double:
   Dim Mzx_B() As Double:       Dim Mzy_B() As Double:

   Dim Mzx_Za() As Double:        Dim Mzy_Za() As Double:
   Dim Mzx_Zb() As Double:        Dim Mzy_Zb() As Double:

' --- Wrap thickness -------------------------------------------------------
'
    Dim phi_i     As Double:    Dim Phi_o     As Double
    Dim phi_iw()  As Double:    Dim phi_ow()  As Double
    Dim gzai_fs() As Double:    Dim gzai_os() As Double

    Dim Wp_xfi As Double:       Dim Wp_yfi As Double
    Dim Wp_xfo As Double:       Dim Wp_yfo As Double
    Dim Wp_xmi As Double:       Dim Wp_ymi As Double
    Dim Wp_xmo As Double:       Dim Wp_ymo As Double

    Dim w_xfi() As Double:      Dim w_yfi() As Double
    Dim w_xfo() As Double:      Dim w_yfo() As Double
    Dim w_xmi() As Double:      Dim w_ymi() As Double
    Dim w_xmo() As Double:      Dim w_ymo() As Double
    Dim thick_fs() As Double:   Dim thick_os() As Double

'
    Dim srt_time_1 As Variant:  Dim end_time_1 As Variant
    Dim srt_time_2 As Variant:  Dim end_time_2 As Variant

' --- Radius of curvature on Wrap envelope -------------------------------------------------------
'
    Dim w_Rcurv_base()  As Double
    Dim w_Rcurv_base_min() As Double
    Dim Wrap_Start_angle_min(4) As Double

    Dim Rcurvature_min_b As Double
    Dim Rcurvature_min_g1 As Double, Rcurvature_min_g2 As Double

    Dim Wrap_Start_angle_min_b As Double
    Dim Wrap_Start_angle_min_g1 As Double, Wrap_Start_angle_min_g2 As Double

'-----------------------------------------------------------------------------------------------
'   Solve Force of Matrix
'-----------------------------------------------------------------------------------------------
    Dim result_0() As Double

      Dim Fk_1 As Double             '  変数1　F1,F2：OSキー溝に働く荷重（OS側Key面の反力） (N)
      Dim Fk_2 As Double             '  変数2　F1,F2：OSキー溝に働く荷重（OS側Key面の反力） (N)
      Dim Fk_3 As Double             '  変数3　F3,F4：MFキー溝に働く荷重（MF側Key面の反力） (N)
      Dim Fk_4 As Double             '  変数4　F3,F4：MFキー溝に働く荷重（MF側Key面の反力） (N)
      Dim Fsb_r As Double             '  変数5　Fsbr：半径方向の軸受け反力(偏心軸、旋回軸) (N)
      Dim Fsb_t As Double             '  変数6　Fsbt：周方向の軸受け反力(偏心軸、旋回軸) (N)
      Dim R_Fsp_t As Double             '  変数7　Ct：FspのOS中心からの距離(鏡板中心z軸からのt方向距離) (mm)
      Dim R_Fsp_r As Double             '  変数8  Cr：FspのOS中心からの距離(鏡板中心からのr方向距離) (mm)

      ' dim kappa as double             '  物性値(一定)断熱指数：κ
      Dim myu_ky As Double             '  物性値(一定)　μk：キー溝の摩擦整数 (ORキー部摩擦係数)
      Dim myu_th As Double             '  物性値(一定)　μt：スラスト面の摩擦係数 (OSスラスト部摩擦係数)
      Dim myu_sb As Double             '  物性値(一定)　μb：旋回軸受け摩擦係数 (OS偏心軸受け部摩擦係数)
Dim atm_00 As Double             '  物性値(一定)atmospheric pressure = 1atm
Dim gravity As Double             '  物性値(一定)gravity(N/sec^2)

      ' dim the_c as double             '  運転条件軸回転角度(圧縮開始=0)　the_c()=the(0)-the() (rad)
      ' dim N_rps as double             '  運転条件運転周波数（rps）
      Dim N_omega As Double             '  運転条件　ω：角加速度(=2πN_rps)
      ' dim P_discharge as double             '  運転条件吐出圧力：Pd (MPaG)
      ' dim P_suction as double             '  運転条件吸入圧力：Ps (MPaG)
      ' dim P_groove as double             '  運転条件スラスト油溝部圧力：Pw (MPaG)
      Dim P_back As Double             '  運転条件背圧室圧力(平均背圧) ：P b (MPaG)

      ' dim Ro as double             '  設計値(定数)旋回半径 (mm)
      Dim R_eb As Double             '  設計値(定数)　rb：旋回軸受け半径 (mm)
      Dim R_eb_out As Double             '  設計値(定数)　OS boss部外径の半径値(ｍｍ)
      Dim L_eb_out As Double             '  設計値(定数)　OS boss部の長さ(ｍｍ)
      Dim m_os As Double             '  設計値(定数)OSの質量 (g)
      Dim vol_os As Double             '  設計値(定数)OSの体積(mm2)
      Dim dense_os As Double             '  設計値(定数)OSの密度 (g/cm3)
      Dim m_or As Double             '  設計値(定数)　m0：Oldham's Ring : ORの質量 (g)
      Dim vol_or As Double             '  設計値(定数)ORの体積 (mm2)
      Dim dense_or As Double             '  設計値(定数)ORの密度 (g/cm3)

'      Dim hw As Double             '  設計値(定数)　h：Wrap高さ(スラスト面からのｚ方向距離) (mm)
      Dim h_pl As Double             '  設計値(定数)　d：OS鏡板厚さ(スラスト面からのｚ方向距離) (mm)
      Dim Z_eb As Double             '  設計値(定数)　L：OSのボス部長さ(スラスト面からのｚ方向距離) (mm)
      Dim Z_mg As Double             '  設計値(定数)  Zm：OSの重心位置(スラスト面からのｚ方向距離) (mm)
      Dim R_Fmg_r As Double             '  設計値(定数)  Wr：OSの重心位置（r寸法）(鏡板中心からのr方向距離) (mm)
      Dim R_Fmg_t As Double             '  設計値(定数)　Wt：OSの重心位置（t寸法）(鏡板中心z軸からのt方向距離) (mm)

      Dim alpha_ky As Double             '  　αky　：キー溝部の設置角度(OS-XY座標系) (rad)
      Dim delta_ky As Double             '  設計値(定数)δ=64.1626degは、θ=0deg(A室吸入完了)時の偏心方向角度とOSｷｰ溝との位相角度 (rad)
      Dim h_ky As Double             '  設計値(定数)　b：キー高さ(スラスト面からのｚ方向距離) (mm)
      Dim L_kcv As Double             '  　Lky　：OS側キー溝部の長さ (mm)
      Dim b_kos As Double             '  設計値(定数)　W1：OS側キー幅(≒溝幅) (mm)
      Dim b_kmf As Double             '  設計値(定数)　 W2：MF側キー幅(≒溝幅) (mm)
      Dim R_kos As Double             '  設計値(定数)　Roy：OS側キー部中心距離(オルダム中心z軸から作用点までの距離) (mm)
      Dim R_kmf As Double             '  設計値(定数)　Rox：MF側キー部中心距離(オルダム中心z軸から作用点までの距離) (mm)

      Dim R_or_out As Double             '  設計値(定数)OldhamRing  Ring部の外径 (mm)
      Dim R_or_in As Double             '  設計値(定数)OldhamRing  Ring部の内径 (mm)
      Dim h_or As Double             '  設計値(定数)OldhamRing  Ring部の高さ (mm)
      Dim L_kos As Double             '  設計値(定数)　OS側キー長さ((mm)
      Dim L_kmf As Double             '  設計値(定数)OS側キー長さ((mm)

      Dim Fgc_r As Double             '  計算(定量値)　Fr：半径方向ガス荷重(ガス力) (N)
      Dim Fgc_t As Double             '  計算(定量値)　Ft：周方向ガス荷重(ガス力) (N)
      Dim Fgc_z As Double             '  計算(定量値)　Fa：圧縮室の軸方向ガス力(ガス力) (N)
      Dim Fgb_z As Double             '  計算(定量値)　Fb：背圧室の軸方向ガス荷重(ガス力) (N)
      Dim Fsp_z As Double             '  計算(定量値)　Fsp：スラスト反力(=Fb-Fa-Fw) (N)
      Dim Fmc_r As Double             '  計算(定量値)　Fc：OSの遠心力 (N)
      Dim Fc_or As Double             '  計算(定量値)　Fc：OldhamRingの遠心力 (N)

      Dim F_mg As Double             '  設計値(定数)  Fw：OSの重量(重力)
      Dim Fs_r As Double             '  推定値　Fs：Wrap同士の押接力(Wrap接点法線≒半径方向)(Wrap反力)　 (N)
      Dim M_sb As Double             '  計算(定量値)　Mb：OS旋回軸受け摩擦モーメント(偏心軸=旋回軸) (Nm)

      Dim R_Fgz_r As Double             '  計算(定量値)　Ar：Faのr座標位置(鏡板中心からのr方向距離) (mm)
      Dim R_Fgz_t As Double             '  計算(定量値)At：FaのOS中心からの距離(鏡板中心z軸からのt方向距離) (mm)
      Dim R_Fgb_r As Double             '  計算(定量値)　Br：Fbのr座標位置(鏡板中心からのr方向距離) (mm)
      Dim R_Fgb_t As Double             '  計算(定量値)　Bt：FbのOS中心からの距離(鏡板中心z軸からのt方向距離) (mm)
      Dim R_Fgc_t As Double             '  計算(定量値)　β：FtのOS円中心からの距離(鏡板中心z軸からのt方向距離) (mm)
      Dim R_Fgc_r As Double             '  計算(定量値)　γ：FrのOS中心からの距離(鏡板中心z軸からのt方向距離) (mm)

   ' 結果比較用
      Dim R_x_osw(4) As Double      ' OS 重心結果　結果比較用
      Dim R_y_osw(4) As Double
      Dim R_z_osw(4) As Double
      Dim V_osw(4) As Double

      Dim R_x_or(4) As Double      ' Oldham Ring 重心結果　結果比較用
      Dim R_y_or(4) As Double
      Dim R_z_or(4) As Double
      Dim V_or(4) As Double

      Dim Ros_F1_oy  As Double '55.969
      Dim Ros_F2_oy  As Double '52.031
      Dim Ros_F3_ox  As Double '57#
      Dim Ros_F4_ox  As Double '7#


   ' 各種結果
      Dim Tilting_os As Double
      Dim Stability_os As Double

      Dim Torque_s As Double
      Dim Fsb_e As Double
      Dim Moment_e As Double
      Dim delta_e As Double
   '
   '


'####################################################################
'
'   for Algebraic Spiral Scroll
'  Referance) Trans. of the JAR Vol.11,No.3(1994)pp.337-347, Hirokatsu KOHSOKABE
'   代数螺線を基本にしたスクロール流体機械のスクロール形状に関する研究
'  Study on Scroll Profile based on Algebraic Spiral for Scroll Fluid Machine
'
'====================================================================
'  Main Routin     ' 2017.9.3-     by shintaku
'====================================================================

Public Sub Alg_Main()

'-------------------------------------------------------------------【M0-1】
'-- Input value to Const  　                        ：定数の設定
'---------------------------------------
    Debug.Print ""
    Debug.Print "■M0     Start time= "; Time      '　vbCrLf
    srt_time_1 = Time

    '----------------------------------------------------
    Call Set_Const

    If k > 1 Then   ' Wrap property
       kk_Flag = -1
       Stop
    Else:     kk_Flag = 1
    End If


'GoTo Label_Main_end


'------------------------------------------------------------------【M0-2】
'-- Input Const's name and value to Alg_Const()     ：定数名の設定
'---------------------------------------
'    Debug.Print "  Input Const's name "
'    Call Set_ConstName                              ' ⇒Alg_Const()

'------------------------------------------------------------------【M1-0】
' 角度の配列番号　：　圧縮開始～終了の角度分割と配列番号を設定
'
'　容積計算１　- Theta ,Phi2 Index Number Set
'【 Calc_Phi2_Index_set 】
'---------------------------------------
'   　圧縮角度範囲と、計算する角度と分割数を決め、昇順に配列に格納する
'       A Chamber Phi_2_Amax --> Phi_2_Amin
'       B Chamber Phi_2_Bmax --> Phi_2_Bmin

'     使用Sub　　Calc_Phi2_AB_Max_to_Min
'         Func   → Phi_2_Amin = Fn_Phi_2(Phi_1_Amin, DataSheetName)
'                → Phi_1_Amax = Fn_Phi_1(Phi_2_Amax, DataSheetName)
'  　 Goalseek使用Cell　：
'        .Range("ZZ1").Value = Phi_2 - 2 * PI                        ' phi2(i)  set Initial value to cell
'        .Range("ZY1").Value = Phi_2 - Atn(k / Phi_2) - 2 * PI       ' set tmporary value of θ to tenporary cell
'        .Range("ZY2").Value = k                                     ' Algebraic constat
'        .Range("AAA1").Formula = "=(ZZ1-atan(ZY2/ZZ1))-ZY1"         ' [Formura No.(12)]
'        .Range("AAA1").GoalSeek Goal:=0, ChangingCell:=Range("ZZ1") ' Goalseek
'---------------------------------------

 Flag_n1 = 0                  '【Flag_n1=1】計算後Cell書出, 【Flag_n1=0】 Cell値の読込利用

   DataSheetName = "DataSheet_5"

   Debug.Print "■M1-1  容積計算1 : Call Calc_Theta_Index_Number_set"

      '-Set Index of The()-----------------------------------------【M1-1】～M1-1-1/-F1/-F2
        Call Calc_Theta_Index_Max_set
            ReDim the(dw_n):            ReDim the_c(dw_n):
            ReDim phi1(dw_n):           ReDim phi2(dw_n)
'            ReDim phi2m1(dw_n)

            ReDim Phi_c_fi(9, dw_n)    ' phi of the contact point from outer to center
'            ReDim Phi_c_fo(9, dw_n)
'            ReDim Phi_c_mi(9, dw_n)
'            ReDim Phi_c_mo(9, dw_n)

      '-Calc The()-------------------------------------------------【M1-2】～M1-2-1
        Call Calc_Theta_Index_Number_set
        Call Calc_Theta_Index_Number_set_check

      '--Get_Radius_Curvature  ------------------------------------【  】
        '   start angel of Wrap head part  g1,g2の計算
        '      使用Sub　：　Get_Radius_curvature_min()
        '      使用Func ：　Fn_Radius_Curvature(Rc_x As Double,  Rc_k As Double)'
        '      未使用  　Fn_Wrap_Start_angle_min(
        '------------------------------------------------------------------
        '   *** wtih parameter=k , start angle of inner and outer wrap
            '   Call Get_Radius_Curvature_parameter

            '--time stamp
              Debug.Print " strat <Get_Radius_Curvature> " & Format(Time, "  HH:mm:ss")

        Call Get_Radius_Curvature     ' paste to DataSheetName = "DataSheet_6"
        'Stop


      '-Calc Phi2(),Phi1()-----------------------------------------【M1-2】～M1-2-1
            '--time stamp
              Debug.Print " strat <Calc_Theta_Index_to_Phi> " & Format(Time, "  HH:mm:ss")

        DataSheetName = "DataSheet_5"

'Call Calc_Theta_Index_to_Phi_test
'Stop

        If Flag_n1 = 0 Then
            ' If Flag_n1 = 0 And (Sheets(DataSheetName).Cells(dw_n + 4, 23).Value <> "") Then
              Debug.Print "   < Read Data from Cell >"

            '-Read Data------------------------------------------【M1-3】
              Call Calc_Theta_Index_Number_Read   ' **from "DataSheet_5"
        Else
            '-Calc Data----------------------------------------- 【M1-4】～M1-4-F1
            '   Phi_c_fi(I, J) =
              Call Calc_Theta_Index_to_Phi        ' ** to "DataSheet_5"
                ' Call Calc_Theta_Index_to_Phi_all_0

        End If

            '--time stamp
              Debug.Print " end   <Calc_Theta_Index_to_Phi> " & Format(Time, "  HH:mm:ss")

      '- Wrap Contact Number of Chamber A and B
        ReDim N_wrap_a(dw_n):            ReDim N_wrap_b(dw_n):
        DataSheetName_2 = "tmp"

        Call Calc_Theta_Index_to_Wrap_Number     ' Number of Wrap contact points at each index

      '-------------------------------------------------------------
      '    Call Calc_Theta_Index_Number_Redim


        '------------------------------------------------------------------【M1-5】
        '　start angel of Wrap head part  g1,g2の計算
        '
        '  Get_Radius_Curvature
        '      使用Sub　：　Get_Radius_curvature_min()
        '      使用Func ：　Fn_Radius_Curvature(Rc_x As Double,  Rc_k As Double)'
        '      未使用  　Fn_Wrap_Start_angle_min(
        '------------------------------------------------------------------

                ' *** wtih parameter=k , start angle of inner and outer wrap
                '  Call Get_Radius_Curvature_parameter
                '  Stop

              'Call Get_Radius_Curvature
              'Stop


'
'    Debug.Print "■M1-5■-容積計算1 経過時間 = " & Format(Time - srt_time_1, "HH:mm:ss")
'    Debug.Print

        '-----------------------------------------------------------【M1-6】
        '　容積計算２　- A、B室容積計算用のφ2、φ1を配列に格納
        '
        '  使用関数　Func Phi_1 = Fn_Phi_1(Phi_2, DataSheetName)
        '  　 Goalseek使用Cell　：
        '        .Range("ZZ1").Value = Phi_2 - 2 * PI                        ' phi2(i)  set Initial value to cell
        '        .Range("ZY1").Value = Phi_2 - Atn(k / Phi_2) - 2 * PI       ' set tmporary value of θ to tenporary cell
        '        .Range("ZY2").Value = k                                     ' Algebraic constat
        '        .Range("AAA1").Formula = "=(ZZ1-atan(ZY2/ZZ1))-ZY1"         ' [Formura No.(12)]
        '        .Range("AAA1").GoalSeek Goal:=0, ChangingCell:=Range("ZZ1") ' Goalseek
        '　　【Flag_n1=1】計算後Cell書出, 【Flag_n1=0】 Cell値読込利用
        '---------------------------------------

        '  Flag_n1 = 0                             '【Flag_n1=1】計算後Cell書出, 【Flag_n1=0】 Cell値の読込利用
        '
        ''    DataSheetName = "DataSheet_5"
        '    '-------------------------------------------------------【M1-6】
        '    Call Calc_Phi2_Phi1
        '
        '    Debug.Print ""
        '    Debug.Print "■M1-6■ 容積計算2 : Call Calc_Phi2_Phi1"
        '    Debug.Print "　経過時間 = " & Format(Time - srt_time_1, "HH:mm:ss")




'■■Label point
Label_Data_Strage_Paste_Label_M2:

'------------------------------------------------------------------【M2-1】
'　容積計算３　- φ2でのA、B室容積を計算
'   【Calc_Phi_to_Volume_A】
'   【Calc_Phi_to_Volume_B】
'      使用Function Phi_1 = Fn_Phi_1(Phi_2, DataSheetName)
'      使用Function V_a(jj_a) = Fn_Calc_Volume_A(Phi_1, Phi_2)
'      使用Function V_b(jj_b) = Fn_Calc_Volume_B(Phi_1, Phi_2)
'------------------------------------------------------------------

    ReDim curve_xw(3, div_n):           ReDim curve_yw(3, div_n):

    '-----------------------------------------------------------
    '   Call Get_xy_Chamber_all           '-- Calculate xy of warp at each index

    '-- Calculate Volume of Chamber A  --------------------------【M2-1】　～M2-1-F1
      Call Calc_Phi_to_Volume_A        '-- Calculate Volume of Chamber A
          ' Index = 0 to dw_n

    '-- Calculate Volume of Chamber B  --------------------------【M2-2】　～M2-2-F1
      Call Calc_Phi_to_Volume_B        '-- Calculate Volume of Chamber B
          ' Index = 0 to dw_n

    '-----------------------------------------------------------
    '    Call Calc_Gravity_Center_chamber

    '-----------------------------------------------------------
    '    Call Calc_Phi_to_Volume_D

        Debug.Print ""
        Debug.Print "■M2■-容積計算3 : Volume A,B"
        Debug.Print "　経過時間 = " & Format(Time - srt_time_1, "HH:mm:ss")

'------------------------------------------------------------
'   Calc_Gravity_Center_wrap
'------------------------------------------------------------

      ' checking the drawing curve
      ReDim curve_xw(11, div_n):          ReDim curve_yw(11, div_n):

      ReDim x_out(div_n):                 ReDim y_out(div_n):
      ReDim x_in(div_n):                  ReDim y_in(div_n):


 '------------------------------------------------------------
  Index_I = 0   ' rotation angle of shaft
                ' Index  0 = 0deg
                ' Index 45 = 90deg
                ' Index 90 = 180deg   , 105 = 207.328deg(A_dis.),  127=249.363deg(B_dis.)
                ' Index 138 = 270deg
                ' Index 183 = 360deg

  '  For Index_i = 0 To 0        ' a revolution of shaft. (rotation angle)

     '---------------------------------------
        Debug.Print "【M2-3】Start / Calc_Gravity_Center_wrap"
        Debug.Print "　経過時間 = " & Format(Time - srt_time_1, "HH:mm:ss")

      Call Calc_Gravity_Center_wrap

        Debug.Print "【M2-3】End / " & Format(Time - srt_time_1, "HH:mm:ss")
      Stop

      Call Calc_Gravity_Center_Mass_OS

        Debug.Print "【M2-4】End / Calc_Gravity_Center_Mass_OS"
        Debug.Print "　経過時間 = " & Format(Time - srt_time_1, "HH:mm:ss")

      Call Calc_Gravity_Center_Mass_OldhamRing

        Debug.Print "【M2-5】End / Calc_Gravity_Center_Mass_OldhamRing"
        Debug.Print "　経過時間 = " & Format(Time - srt_time_1, "HH:mm:ss")

  '  Next Index_i



'■■Label point
Label_Data_Strage_Paste_Label_M3:
'------------------------------------------------------------------【M3-1】
'       圧力計算
'---------------------------------------

      Call Calc_Press_AB

        Debug.Print "■M3■ -圧力計算 Pressure A,B"
        Debug.Print "　処理時間 = " & Format(Time - srt_time_1, "HH:mm:ss")



'■■Label point
Label_Data_Strage_Paste_Label_M6:

'------------------------------------------------------------------【M6-1】
'　OS作用力計算
'　　　Ft : Calc_GasForce_Ft　tangential 接線方向
'　　　Fr : Calc_GasForce_Ft  Radial direction 接線方向
'
'---------------------------------------

 Index_I = Index_I

      ' Index  0 = 0deg
      ' Index 45 = 90deg
      ' Index 90 = 180deg   , 105 = 207.328deg(A_discharge),  127=249.363deg(B_discharge)
      ' Index 138 = 270deg
      ' Index 183 = 360deg


    Call Calc_GasForce_Ft
      '   Ft  : Fgc_t       [ R_Fgc_t ]               '[N] , [mm]
      '   Ft  : Ft_AB(i,j)  [ Lt_AB(i, j) ]           '[N] , [mm]
      '        Ft_A  : Ft_A(i,j)  [ LLt_A(i,j) ]
      '        Ft_B  : Ft_B(i,j)  [ LLt_B(i,j) ]
      '        Ft_D  : Ft_D(i,j)  [ Lt_D(i,j) ]
      '   Mt  : Mmt_AB(i,j)                           ' [Nm]  <--単位に注意
      '        Mmt_A  : Mmt_A(i,j)  [ Lmt_A(i, j)]
      '        Mmt_B  : Mmt_B(i,j)  [ Lmt_B(i, j)]
      '        Mmt_D  : Mmt_D(i,j)  [ Lmt_D(i, j)]
      '


    Call Calc_GasForce_Fr
      '   Fr  : Fgc_r       [ R_Fgc_r ]               '[N] , [mm]
      '   Fr  : Fr_AB(i,j)  [ Lr_AB(i, j) ]           '[N] , [mm]
      '        Fr_A  : Fr_A(i,j)  [ LLr_A(i,j) ]
      '        Fr_B  : Fr_B(i,j)  [ LLr_B(i,j) ]
      '        Fr_D  : Fr_D(i,j)  [ Lr_D(i,j) ]
      '   Mr  : Mmr_AB(i,j)                           ' [Nm]  <--単位に注意
      '        Mmr_A  : Mmr_A(i,j)  [ Lmr_A(i, j)]
      '        Mmr_B  : Mmr_B(i,j)  [ Lmr_B(i, j)]
      '        Mmr_D  : Mmr_D(i,j)  [ Lmr_D(i, j)]

    Call Calc_GasForce_Fz       '20180606

      '   Fa  : Fgc_z
      '         Fz_Za(i)  [ x_zforce_a(i) , y_zforce_a(i) ]   '[N] , [mm]

      '   Fb  : Fgb_z
      '         Fz_Zb(i)  [ x_zforce_b(i) . y_zforce_b(i) ]

      '   Fsp : Fsp_z
      '         Fz_sp(i)


    Debug.Print ""
    Debug.Print "■M6-1■ Data保存処理　Data_Strage and Paste"
    Debug.Print "　処理時間 = " & Format(Time - srt_time_1, "HH:mm:ss")


'GoTo label_Data_Strage_Paste_end

    '---------------------------------------
    '　 Matrix term
        Call Get_Matrix_A_and_C             '20180614
    '---------------------------------------
    '　 Calculation of Matrix
        Call Get_Matrix_X

        Call Get_Matrix_and_results



'■■Label point
label_Data_Strage_Paste_end:

    '---------------------------------------
    ' *** aiba結果と照合、検証 ***
    '---------------------------------------

          DataSheetName_3 = "tmp"
          ' DataSheetName_3 = "DataSheet_D2"

          Call Data_Strage_to_array_2           ' aiba結果と照合、検証用
            ' Call Data_Strage_to_array_3       ' 追加：接点座標、角度、動径

'Stop
    '---------------------------------------
    ' *** 結果検証 ***
    '---------------------------------------
          DataSheetName = "DataSheet_6"
          Call Data_Strage_to_array             ' 追加：接点座標、角度、動径

          Debug.Print "■ 計算 END time= "; Time
          Debug.Print "　処理時間 = " & Format(Time - srt_time_1, "HH:mm:ss")

    '---------------------------------------
    ' *** 表示Sheet ***
    '---------------------------------------
          DataSheetName_3 = "tmp"
          Sheets(DataSheetName_3).Select


Stop


'GoTo Label_Main_end



'■■Label point
Label_Data_Strage_Paste_Label_M4:

'------------------------------------------------------------------【M4-1】
'　Data用の配列に読込み、Cellに書出
'【Data_Strage_to_array】
'---------------------------------------



'■■Label point
Label_Data_Strage_Paste_Label_M5:

'------------------------------------------------------------------【M5-1】
'　Wrap厚計算
'【Calc_Wrap_thickness】
'       使用Sub　 Call Calc_Solve_thickness
'       使用Sub　Call Wp_xyfi(Phi_i), Call Wp_xyfo(Phi_o)
'       使用Sub　Call Wp_xymi(Phi_o), Call Wp_xymo(Phi_o)
'
'  　 Solver使用Cell　：
'        .Range("ZZ1").Value = Phi_o                         ' Phi_o : φ2out Initial value
'        .Range("ZY1").Value = Phi_i                         ' Phi_i : φ2in
'        .Range("ZY2").Value = k                             ' Algebraic constat
'        .Range("ZY3").Value = a                             ' Algebraic constat
'
'        .Range("AAA2").Formula = "ξ’の分子/分母
'        .Range("AAA4").Formula = "=ZY3*(ZY1^ZY2*sin(ZY1)+ZZ1^ZY2*sin(ZZ1))
'        .Range("AAA5").Formula = "=sin(ZY1-atan(ZY2/ZY1))-sin(ZZ1-atan(ZY2/ZZ1))"
'        .Range("AAA1").Formula = "=(AAA4+(AAA2)*AAA5)^2"
'
'---------------------------------------
'　分割 tmp_n0=15 (deg)　 ⇒ etc.　phi2(i) = OS_in_end - (tmp_n0 / 180 * PI) * i


         Flag_n1 = 0        '【Flag_n1=1】計算後Cell書出, 【Flag_n1=0】 Cell値読込利用

            DataSheetName = "DataSheet_4"

            '------------------------------------------------------- 【M5-1】
            Call Calc_Wrap_thickness               '-- Calculate Volume of Chamber A

            Debug.Print ""
            Debug.Print "■M5■ Wrap厚　Call Calc_Wrap_thickness"
            Debug.Print "　処理時間 = " & Format(Time - srt_time_1, "HH:mm:ss")

 'Stop


        '    '-- Caculate Scroll Wrap angle                      ：代数螺旋、包絡線の偏角
        '    Debug.Print "  Caculate Wrap angles" & vbCrLf
        '        Call Calc_WrapAngle
        '
        '    '-- Caculate Scroll Wrap XY                         ：包絡線の座標
        '    Debug.Print "  Caculate Scroll Wrap XY" & vbCrLf
        '        Call Calc_Wrap_XY




'■■Label point
'Label_Main_end:


Stop


End Sub


'======================================================== 【M0-1】
'　定数の設定
'　     代数螺旋のParameter　a,k,φ,β,θc
'       包絡線のParameter    ξ1,ξ2,ψ
'========================================================

Public Sub Set_Const()
Dim set_i As Long

'-- Input value to Const
'
    pi = Atn(1) * 4
    dw_c = 5                    ' Data_Strage 配列の列数-初期値

    dw_deg = 2                  ' [deg] Division width（>0)
    dw = dw_deg * pi / 180      ' [rad] Division width（>0)
    div_n = 360                 ' Chanber面積の分割数　台形公式利用
    dw_t_deg = 15               ' [deg] Wrap thckness Division width（>0)
    dw_t = dw_t_deg * pi / 180  ' [rad] Wrap thckness Division width（>0)

'    DataSheetName = "DataSheet_3"

    kappa = 1.07765         ' R401a
    P_suction = 0.9978      ' [MPa(abs)] Suction Pressure
    P_discharge = 3.3853    ' [MPa(abs)] Discharge Pressure

    a = 3.8                 ' Algebraic constat     ("C6"=cell(6,3) )
    k = 0.91                ' Algebraic constat

'    g1 = 1.46   '  1.46     ' Algebraic constat     g1 = 2.26
'    g2 = 3.06   '  3.06     ' Algebraic constat     g2 = 2.26
'        g1 = 3.06     ' Algebraic constat     g1 = 2.26
'        g2 = 1.46     ' Algebraic constat     g2 = 2.26
          g1 = 2.26
          g2 = 2.26

    qq = 1.8325956          ' Axis Rotation angel   qq = 0
    Ro = 4.52               ' Obit radius
    Hw = 38.835                             '[mm]  "Wrap Hight [mm]"

    FS_in_srt_deg = 142             '[deg] "FS in srt"
    FS_in_end_deg = 1052          '[deg] "FS in end"
    FS_out_srt_deg = 106            '[deg] "FS out srt"
    FS_out_end_deg = 872           '[deg] "FS out end"

    OS_in_srt_deg = 106             '[deg] "OS in srt"
    OS_in_end_deg = 872             '[deg] "OS in end"
    OS_out_srt_deg = 142            '[deg] "OS out srt"
    OS_out_end_deg = 1052           '[deg] "OS out end"

    FS_in_srt = FS_in_srt_deg * pi / 180    '[rad] "FS in srt"
    FS_in_end = FS_in_end_deg * pi / 180    '[rad] "FS in end"
    FS_out_srt = FS_out_srt_deg * pi / 180  '[rad] "FS out srt"
    FS_out_end = FS_out_end_deg * pi / 180  '[rad] "FS out end"

    OS_in_srt = OS_in_srt_deg * pi / 180    '[rad] "OS in srt"
    OS_in_end = OS_in_end_deg * pi / 180    '[rad] "OS in end"
    OS_out_srt = OS_out_srt_deg * pi / 180  '[rad] "OS out srt"
    OS_out_end = OS_out_end_deg * pi / 180  '[rad] "OS out end"

    OS_offset_angle = 0 * pi / 180          '[rad] "OS Wrap offset angle"
    FS_offset_angle = 0 * pi / 180          '[rad] "FS Wrap offset angle"
    dx = 3.7         ' = 3.7           '[mm]  "FS Wrap offset dx"
    dy = 2.4         ' = 2.4           '[mm]  "FS Wrap offset dy"
    RR_dxy = Sqr(dx ^ 2 + dy ^ 2)      '[mm]  "offset radius
    alpha_dxy = Atn(dy / dx)           '[rad] "offset direction

'------------------------------------------------
' OS Plate
    OS_dia = 125#        '[mm]  "OS diameter"
    OS_seal = 62#        '[mm]  "OS seal outer diameter"

'------------------------------------------------
'  Suction inlet area shape
'  1) inlet FS inner line : radius = 57.14  Center (xc,yc)=(4.47,3.32)
'  2) inlet FS outer line : radius = 44.1  Center (xc,yc)=(6.02,1.2)
'------------------------------------------------
      r_Rfi_c = 57.14          ' 57.14
      x_Rfi_c = 4.47
      y_Rfi_c = 3.32
      angle_Rfi_c = (360 - 89) * pi / 180     ' on FS Drawing xy-cordinate

      r_Rfo_c = 44.1
      x_Rfo_c = 6.02
      y_Rfo_c = 1.2
      angle_Rfo_c = (360 - 89) * pi / 180     'on FS Drawing xy-cordinate

'---------------
' Wrap head arc dimension
'---------------
 'FS-inner
  set_i = 1
      R_head_r1(set_i) = 6.5
      R_head_xc(set_i) = -6.5921:       R_head_yc(set_i) = 6.3337
        R_head_x1(set_i) = -3.2201:       R_head_y1(set_i) = 0.7769
        R_head_x2(set_i) = -12.7945:      R_head_y2(set_i) = 8.2779
 'FS-ouer
  set_i = 2
      R_head_r1(set_i) = 1#
      R_head_xc(set_i) = -0.0039:       R_head_yc(set_i) = 0.9059
        R_head_x1(set_i) = -0.1776:       R_head_y1(set_i) = 1.8907
        R_head_x2(set_i) = 0.9054:        R_head_y2(set_i) = 1.322

 'OS-inner
  set_i = 3
      R_head_r1(set_i) = 6.5
      R_head_xc(set_i) = 0.9143:        R_head_yc(set_i) = 0.5396
        R_head_x1(set_i) = 3.646:         R_head_y1(set_i) = 6.4378
        R_head_x2(set_i) = -4.9955:       R_head_y2(set_i) = 3.2459
 'OS-ouer
  set_i = 4
      R_head_r1(set_i) = 1#
      R_head_xc(set_i) = 7.514:         R_head_yc(set_i) = 6.6698
        R_head_x1(set_i) = 7.514:         R_head_y1(set_i) = 5.6698
        R_head_x2(set_i) = 8.4683:        R_head_y2(set_i) = 6.9686


'------------------------------------------------
' Oil Groove
'   arc1, arc2 spec
'------------------------------------------------
  '-- arc1 spec.
   r1_oilgroove = 53.72           'outer radius 53.72  / inner radius=52.22
   x1_oilgroove_c = 0.89
   y1_oilgroove_c = 2.4
   t1_oilgroove = 1.5
   angle1_oilgroove_0 = (-70) * pi / 180         ' start angle  'on arc center

  '-- arc2 spec.  near by suction inlet
   r2_oilgroove = 56.35           'outer radius 56.35  / inner radius=54.85
   x2_oilgroove_c = 1.05
   y2_oilgroove_c = -0.25
   t2_oilgroove = t1_oilgroove
   angle2_oilgroove_2 = (180 + 70) * pi / 180    ' end angle    " In case of Mr.aiba (180 + 70)"
'   angle2_oilgroove_2 = (180 + 61) * pi / 180     ' end angle    '= (180 + 62.5)


'-----------------------------------------------------------------------------------------------
'   Solve Force of Matrix
''-----------------------------------------------------------------------------------------------
' ' 初期値、定数  設定
'    Fk_1 = 0
'    Fk_2 = 0
'    Fk_3 = 0
'    Fk_4 = 0
'    Fsb_r = 0
'    Fsb_t = 0
'    R_Fsp_t = 0
'    R_Fsp_r = 0

    kappa = 1.08
    myu_ky = 0.02
    myu_th = 0.04
    myu_sb = 0.01
    atm_00 = 0.1013   ' for Gage pressur
    gravity = 9.80665


   ' the_c =0
   N_rps = 58
   ' N_omega =0
'   P_discharge = 3.284 + atm_00              ' [MPa(abs)]
'   P_suction = 0.896 + atm_00                ' [MPa(abs)]
   P_groove = 1.84541769181162 + atm_00      ' [MPa(abs)]
   P_back = 1.5464 + atm_00                  ' [MPa(abs)]

   Ro = 4.52
   R_eb = 28 / 2
   R_eb_out = 21
   L_eb_out = 31.8
'   m_os = 1485.35447277754
'   vol_os = 203473.2154
   dense_os = 7.3
'   m_or = 73.93932
'   vol_or = 27084
   dense_or = 2.73

'   h_w = 38.83
   h_pl = 9
   Z_eb = 31.8

'   Z_mg = 1.73403418964931      ' Z_mg      ' < Calc_Gravity_Center_Mass_OS >
'   R_Fmg_r = -1.2116339585943   ' R_Fmg_r   ' < Calc_Gravity_Center_Mass_OS >
'   R_Fmg_t = 1.49259786115127   ' R_Fmg_t   ' < Calc_Gravity_Center_Mass_OS >

   alpha_ky = 20 * pi / 180
   delta_ky = (64.1626) * pi / 180          '[20180719] = (-64.1626) * pi / 180
   h_ky = 6
   L_kcv = 20
   b_kos = 8
   b_kmf = 8.5
   R_kos = 54                   ' key distance of OS side from OS center
   R_kmf = 57:                   ' key distance of MF side from OS center

   R_or_out = 59
   R_or_in = 51
   h_or = 8.3                   '  ring hight of Oldham-ring
   L_kos = 10                   '  key lengrh of OS side
   L_kmf = 12                   '  key lengrh of MF side

   Fs_r = 0                     '  Wrap contact force
   M_sb = 0                     '  Momet of eccentric bearing

'   Fgc_r = 524.05268175235      ' Fr_AB(j)      *** VBA結果 258.74 ***約半分
'   Fgc_t = 4225.06655933097     ' Ft_AB(j)
'   Fgc_z = 18269.0094291992     ' Fz_Za(i)
'   Fgb_z = 24223.165717017      ' Fz_Zb(i)
'   Fsp_z = 5939.59008491286     ' Fz_sp(i)      *** VBA結果
'   Fmc_r = 891.431904510562     ' Fmc_r   < Calc_Gravity_Center_Mass_OS >
'   F_mg = 14.5662029050165      ' F_mg    < Calc_Gravity_Center_Mass_OS >
'
'
'   R_Fgz_r = -0.59094778            '
'   R_Fgz_t = -0.107573854           '
'   R_Fgb_r = -0.978663241909462     '
'   R_Fgb_t = 3.9305633837589E-16    '
'   R_Fgc_t = 5.0072963597173       ' Lt_AB(j)  < Call Calc_GasForce_Ft >
'   R_Fgc_r = -0.824002719804529    ' Lr_AB(j)      *** VBA結果 -0.2684

'  delta_e = 0.1

End Sub



'======================================================== 【M1-1】
'  Caculate Theta θ rotation Angle and Index ()　:
'
'
'    使用Sub　： Calc_Phi2_AB_Max_to_Min　← Goalseek 利用
'　  ・
'      　代数螺旋の偏角φ2､φ1を求める｡ また、設定角度幅毎のφ2とφ1を求める｡
'　  ・ 代数螺旋の偏角φ1、φ2に対する、軸角度θcを求める。
'
'========================================================

Public Sub Calc_Theta_Index_Max_set()            '　20171030

Dim I As Long, J As Long


'--------------------------------------------
' [00]  A,B室の容積Max,Min時のφ2を求め、昇順に列挙
'   Goalseek 利用　φ1 →φ2
'--------------------------------------------

    Call Calc_Phi2_AB_Max_to_Min

'    Debug.Print "[00] P2_C1 =" & Format(P2_C1, " 00.000 / ") & (P2_C1 / pi * 180) & "deg "
'    Debug.Print "     P2_C2 =" & Format(P2_C2, " 00.000 / ") & (P2_C2 / pi * 180) & "deg "
'    Debug.Print "     P2_C3 =" & Format(P2_C3, " 00.000 / ") & (P2_C3 / pi * 180) & "deg "
'    Debug.Print "     P2_C4 =" & Format(P2_C4, " 00.000 / ") & (P2_C4 / pi * 180) & "deg "


'--------------------------------------------
' [0] Theta θ： The() 配列数を決める
'     軸回転角 θ
'　　　A室圧縮開始角度～AorB室最終側の圧縮終了角度に対応する、θcを求める
'      圧縮開始を基点とし、既定角度dw毎の点と、
'　　　　　　　Wrapの開始角C1 , C2､終了角C3､C4点も加える｡
'
'--------------------------------------------

'    Phi_Max = Application.Max(FS_in_srt, FS_in_end, OS_in_srt, OS_in_end)
'    Phi_Min = Application.Min(FS_in_srt, FS_in_end, OS_in_srt, OS_in_end)

'０巻内側　 FS,OS Wrap angle of start & end angle

    The_C1 = P2_C1 - Atn(k / P2_C1)
    The_C2 = P2_C2 - Atn(k / P2_C2)
    The_C3 = P2_C3 - Atn(k / P2_C3)
    The_C4 = P2_C4 - Atn(k / P2_C4)

    The_Max = The_C1
    The_Min = The_C4

' Wrapの巻き数  **Int関数は引数numを超えない最大の負の整数を返す**
    turn_wrap_n = Int((The_Max - The_Min) / (2 * pi)) + 1
    '    turn_wrap_n = Application.RoundUp(FS_in_end / (2 * PI), 0)

    dw_n_end = Int((The_Max - The_Min) / dw) + 1          ' 仮数

    dw_n = Int(turn_wrap_n * 2 * pi / dw) + 1             '171101 条件追加

    I = 3 * turn_wrap_n                                   '171101 追加の配列数　C2,C3,C4
    dw_n = dw_n + I


' --- Wrap angle Theta Region　　1,2巻内側                          '171101
    ReDim The_C1n(turn_wrap_n):     ReDim The_C2n(turn_wrap_n)       '171101
    ReDim The_C3n(turn_wrap_n):     ReDim The_C4n(turn_wrap_n)       '171101
        ReDim dw_n_C1n(turn_wrap_n):       ReDim dw_n_C2n(turn_wrap_n)
        ReDim dw_n_C3n(turn_wrap_n):       ReDim dw_n_C4n(turn_wrap_n)

' j巻内側  　　'171101


' The_C1n(0)

      If The_C1 > (4 * pi) Then
        The_C1n(0) = The_C1
        The_C1n(1) = The_C1 - (2 * pi)
        The_C1n(2) = The_C1 - (4 * pi)
      ElseIf The_C1 > (2 * pi) Then
        The_C1n(1) = The_C1 - (2 * pi)
        The_C1n(2) = The_C1 - (4 * pi)
      End If

' The_C2n(0)

      If The_C2 > (4 * pi) Then
        The_C2n(0) = The_C2
        The_C2n(1) = The_C2 - (2 * pi)
        The_C2n(2) = The_C2 - (4 * pi)
      ElseIf The_C2 > (2 * pi) Then
        The_C2n(0) = The_C2
        The_C2n(1) = The_C2 - (2 * pi)
        The_C2n(2) = The_C2 - (4 * pi)
      End If


'  The_C3n(0)

      If The_C3 > (4 * pi) Then
        The_C3n(0) = The_C3
        The_C3n(1) = The_C3 - (2 * pi)
        The_C3n(2) = The_C3 - (4 * pi)
      ElseIf The_C3 > (2 * pi) Then
        The_C3n(0) = The_C3 + (2 * pi)
        The_C3n(1) = The_C3
        The_C3n(2) = The_C3 - (2 * pi)
      ElseIf (The_C3 < (2 * pi)) And (The_C3 > 0) Then
        The_C3n(0) = The_C3 + (4 * pi)
        The_C3n(1) = The_C3 + (2 * pi)
        The_C3n(2) = The_C3
      End If

' The_C4n(0)

      If The_C4 > (4 * pi) Then
        The_C4n(0) = The_C4
        The_C4n(1) = The_C4 - (2 * pi)
        The_C4n(2) = The_C4 - (4 * pi)
      ElseIf The_C4 > (2 * pi) Then
        The_C4n(0) = The_C4 + (2 * pi)
        The_C4n(1) = The_C4
        The_C4n(2) = The_C4 - (2 * pi)
      ElseIf (The_C4 < (2 * pi)) And (The_C4 > 0) Then
        The_C4n(0) = The_C4 + (4 * pi)
        The_C4n(1) = The_C4 + (2 * pi)
        The_C4n(2) = The_C4
      End If

' （－）符号処理

   For J = 0 To turn_wrap_n
        If The_C1n(J) < 0 Then
            The_C1n(J) = 0
        End If
        If The_C2n(J) < 0 Then
            The_C2n(J) = 0
        End If
        If The_C3n(J) < 0 Then
            The_C3n(J) = 0
        End If
        If The_C4n(J) < 0 Then
            The_C4n(J) = 0
        End If
   Next


'   For J = 0 To turn_wrap_n
'
'      Debug.Print "  The_C1n(" & J & ") =" & Format(The_C1n(J), " 00.000 / ") & (The_C1n(J) * 180 / pi) & "deg"
'      Debug.Print "  The_C2n(" & J & ") =" & Format(The_C2n(J), " 00.000 / ") & (The_C2n(J) * 180 / pi) & "deg"
'      Debug.Print "  The_C3n(" & J & ") =" & Format(The_C3n(J), " 00.000 / ") & (The_C3n(J) * 180 / pi) & "deg"
'      Debug.Print "  The_C4n(" & J & ") =" & Format(The_C4n(J), " 00.000 / ") & (The_C4n(J) * 180 / pi) & "deg"
'
'   Next J

      Debug.Print " "


 End Sub



'======================================================== <M1-1-1>
'　代数螺旋 AB内壁の開始・終了の偏角φ2 の Max,Min Index順番
'　　 Goalseek 利用　A,B室の容積Max,Min時のφ2を求め順番に列挙
'        P1,P2：圧縮室接点の、巻始側(Phi_1)、巻終側(Phi_2)の偏角
'
'     Angle[rad]  Index[i]
'       P2_C1 -  dw_n_C1     : A Max
'       P2_C2 -  dw_n_C2     : B Max
'       P2_C3 -  dw_n_C3     : A or B
'       P2_C4 -  dw_n_C4     : A or B min
'
'           Phi_1_Amax  Phi_1_Amin
'           Phi_2_Amax  Phi_2_Amin
'           Phi_1_Bmax  Phi_1_Bmin
'           Phi_2_Bmax  Phi_2_Bmin
'
'========================================================

Public Sub Calc_Phi2_AB_Max_to_Min()

'---------------------------------------
'         FS in 142 - 1052
'         OS in 106 - 872 (=1052-180)
'---------------------------------------
'        DataSheetName = "DataSheet_3"

' A Chamber Max,Min angle of Base Spiral
'  A室 最内圧縮室の接点　最小偏角
    If FS_in_srt >= OS_out_srt Then
        Phi_1_Amin = FS_in_srt
    Else
        Phi_1_Amin = OS_out_srt
    End If
        Phi_2_Amin = Fn_Phi_2(Phi_1_Amin, DataSheetName)      ' Goalseek用のDatasheet名が必要

'  A室 最大圧縮室の接点　最大偏角
    If FS_in_end <= OS_out_end Then
       Phi_2_Amax = FS_in_end
    Else
       Phi_2_Amax = OS_out_end
    End If
        Phi_1_Amax = Fn_Phi_1(Phi_2_Amax, DataSheetName)

            Phi_2_Amin_deg = Phi_2_Amin * 180 / pi
            Phi_2_Amax_deg = FS_in_end_deg

' B Chamber Max,Min angle of Base Spiral

'  B室 最小圧縮室の接点　最小偏角
   If OS_in_srt >= FS_out_srt Then
        Phi_1_Bmin = OS_in_srt
    Else
        Phi_1_Bmin = FS_out_srt
    End If
        Phi_2_Bmin = Fn_Phi_2(Phi_1_Bmin, DataSheetName)      ' Goalseek用のDatasheet名が必要

    '  B室 最大圧縮室の接点　最大偏角
     If OS_in_end <= FS_out_end Then
        Phi_2_Bmax = OS_in_end
     Else
        Phi_2_Bmax = FS_out_end
     End If
        Phi_1_Bmax = Fn_Phi_1(Phi_2_Bmax, DataSheetName)

            Phi_2_Bmin_deg = Phi_2_Bmin * 180 / pi
            Phi_2_Bmax_deg = OS_in_end_deg
'--
'　　　代数螺旋偏角：　A,B室の圧縮開始と終了時の偏角を、降順列挙
                P2_C1_deg = Phi_2_Amax_deg
                P2_C2_deg = Phi_2_Bmax_deg
                P2_C1 = Phi_2_Amax
                P2_C2 = Phi_2_Bmax

        If Phi_2_Amin_deg > Phi_2_Bmin_deg Then
                P2_C3_deg = Phi_2_Amin_deg
                P2_C4_deg = Phi_2_Bmin_deg
                P2_C3 = Phi_2_Amin
                P2_C4 = Phi_2_Bmin
        Else
                P2_C3_deg = Phi_2_Bmin_deg
                P2_C4_deg = Phi_2_Amin_deg
                P2_C3 = Phi_2_Bmin
                P2_C4 = Phi_2_Amin
        End If


End Sub


'======================================================== <M1-1-F1>
'　関数：代数螺旋の偏角φ1からφ2を求める
'
'
'========================================================

Public Function Fn_Phi_2(Phi_1 As Double, ByVal DataSheetName As String) As Double

    Dim I As Long, J As Long

    Sheets(DataSheetName).Activate

'-----------------
'　　Gaolseek φ1 → φ2
'-----------------
    With Sheets(DataSheetName)
        .Range("ZZ1").Value = Phi_1 + 2 * pi                        ' phi2(i)  set Initial value to cell
        .Range("ZY1").Value = Phi_1 - Atn(k / Phi_1)                ' set tmporary value of θ to tenporary cell
        .Range("ZY2").Value = k                                     ' Algebraic constat
        .Range("AAA1").Formula = "=ZZ1-atan(ZY2/ZZ1)-(ZY1+2*PI())"  ' [Formura No.(12)]
        .Range("AAA1").GoalSeek Goal:=0, ChangingCell:=Range("ZZ1") ' Goalseek

    '      Range("AAA2").Value = phi2(i) - Atn(k / phi2(i)) - (the(i) + 2 * PI)  'φ2　検算表示
    End With

'-----------------
'   Goalseekの結果：φ2 を戻り値に設定
'-----------------
     Fn_Phi_2 = Range("ZZ1").Value


End Function


'Fn1=================================================== 【M5-1-2】 For Wrap thickness
   Public Function Fn_xfi(phi_i As Double)   ' Formura (8) fi
        If phi_i <= 0 Then
           Fn_xfi = 0
'           Fn_xfi = a * phi_i ^ k * Cos(phi_i - qq) + g1 * Cos(phi_i - qq - pi / 2) '+ dx
        Else
           Fn_xfi = a * phi_i ^ k * Cos(phi_i - qq) + g1 * Cos(phi_i - qq - Atn(k / phi_i)) '+ dx
        End If
   End Function

   Public Function Fn_yfi(phi_i As Double)   ' Formura (8) fi
        If phi_i <= 0 Then
           Fn_yfi = 0
'           Fn_yfi = a * phi_i ^ k * Sin(phi_i - qq) + g1 * Sin(phi_i - qq - pi / 2) '+ dy
        Else
           Fn_yfi = a * phi_i ^ k * Sin(phi_i - qq) + g1 * Sin(phi_i - qq - Atn(k / phi_i)) '+ dy
        End If
   End Function


'Fn2=================================================== 【M5-1-3】
   Public Function Fn_xfo(Phi_o As Double)   ' Formura (7) fo
        If Phi_o <= 0 Then
           Fn_xfo = 0
'           Fn_xfo = -a * Phi_o ^ k * Cos(Phi_o - qq) + g1 * Cos(Phi_o - qq - pi / 2) '+ dx
        Else
           Fn_xfo = -a * Phi_o ^ k * Cos(Phi_o - qq) + g1 * Cos(Phi_o - qq - Atn(k / Phi_o)) '+ dx
        End If
   End Function

   Public Function Fn_yfo(Phi_o As Double)   ' Formura (7) fo
        If Phi_o <= 0 Then
           Fn_yfo = 0
'           Fn_yfo = -a * Phi_o ^ k * Sin(Phi_o - qq) + g1 * Sin(Phi_o - qq - pi / 2) '+ dy
        Else
           Fn_yfo = -a * Phi_o ^ k * Sin(Phi_o - qq) + g1 * Sin(Phi_o - qq - Atn(k / Phi_o)) '+ dy
        End If
   End Function


'Fn3=================================================== 【M5-1-4】
   Public Function Fn_xmo(phi_i As Double)   ' Formura (5) mo   '(-21)
        If phi_i <= 0 Then
           Fn_xmo = 0
'           Fn_xmo = a * phi_i ^ k * Cos(phi_i - qq) - g2 * Cos(phi_i - qq - pi / 2) '+ dx
        Else
           Fn_xmo = a * phi_i ^ k * Cos(phi_i - qq) - g2 * Cos(phi_i - qq - Atn(k / phi_i)) '+ dx
        End If
   End Function

   Public Function Fn_ymo(phi_i As Double)   ' Formura (5) mo   '(-21)
        If phi_i <= 0 Then
           Fn_ymo = 0
'           Fn_ymo = a * phi_i ^ k * Sin(phi_i - qq) - g2 * Sin(phi_i - qq - pi / 2) '+ dy
        Else
           Fn_ymo = a * phi_i ^ k * Sin(phi_i - qq) - g2 * Sin(phi_i - qq - Atn(k / phi_i)) '+ dy
        End If
   End Function


'Fn4=================================================== 【M5-1-5】
   Public Function Fn_xmi(Phi_o As Double)   ' Formura (6) mi   '(-14)
        If Phi_o <= 0 Then
           Fn_xmi = 0
'           Fn_xmi = -a * Phi_o ^ k * Cos(Phi_o - qq) - g2 * Cos(Phi_o - qq - pi / 2) '+ dx
        Else
           Fn_xmi = -a * Phi_o ^ k * Cos(Phi_o - qq) - g2 * Cos(Phi_o - qq - Atn(k / Phi_o)) '+ dx
        End If
   End Function

   Public Function Fn_ymi(Phi_o As Double)   ' Formura (6) mi   '(-14)
        If Phi_o <= 0 Then
           Fn_ymi = 0
'           Fn_ymi = -a * Phi_o ^ k * Sin(Phi_o - qq) - g2 * Sin(Phi_o - qq - pi / 2) '+ dy
        Else
           Fn_ymi = -a * Phi_o ^ k * Sin(Phi_o - qq) - g2 * Sin(Phi_o - qq - Atn(k / Phi_o)) '+ dy
        End If
   End Function




'======================================================== <M1-1-F2>
'　関数：　代数螺旋の偏角φ2からφ1を求める
'
'
'========================================================

Public Function Fn_Phi_1(Phi_2 As Double, ByVal DataSheetName As String) As Double

    Dim I As Long, J As Long

    Sheets(DataSheetName).Activate

'-----------------
'　　Gaolseek φ2 → φ1
'-----------------
    With Sheets(DataSheetName)
        .Range("ZZ1").Value = Phi_2 - 2 * pi                        ' phi2(i)  set Initial value to cell
        .Range("ZY1").Value = Phi_2 - Atn(k / Phi_2) - 2 * pi       ' set tmporary value of θ to tenporary cell
        .Range("ZY2").Value = k                                     ' Algebraic constat
        .Range("AAA1").Formula = "=(ZZ1-atan(ZY2/ZZ1))-ZY1"         ' [Formura No.(12)]

        .Range("AAA1").GoalSeek Goal:=0, ChangingCell:=Range("ZZ1") ' Goalseek

    End With

'-----------------
'   Goalseekの結果：φ2 を戻り値に設定
'-----------------
     Fn_Phi_1 = Range("ZZ1").Value


End Function

'======================================================== <M1-1-F2>
'　関数：　代数螺旋の偏角φ2からφ1を求める
'
'
'========================================================

Public Function Fn_Phi_2m1(Phi_2 As Double, ByVal DataSheetName As String) As Double

    Dim I As Long, J As Long

    Sheets(DataSheetName).Activate

'-----------------
'　　Gaolseek φ2 → φ2m1 counter point between φ2 and φ1  (reverse wall contact angle)
'-----------------
    With Sheets(DataSheetName)
        .Range("ZZ1").Value = Phi_2 - pi                         ' phi2(i)  set Initial value to cell
        .Range("ZY1").Value = Phi_2 - Atn(k / Phi_2) - pi        ' set tmporary value of θ to tenporary cell
        .Range("ZY2").Value = k                                     ' Algebraic constat
        .Range("AAA1").Formula = "=(ZZ1-atan(ZY2/ZZ1))-ZY1"         ' [Formura No.(12)]
        .Range("AAA1").GoalSeek Goal:=0, ChangingCell:=Range("ZZ1") ' Goalseek

    '      Range("AAA2").Value = phi2(i) - Atn(k / phi2(i)) - (the(i) + 2 * PI)  'φ2　検算表示
    End With

'-----------------
'   Goalseekの結果：φ2 を戻り値に設定
'-----------------
     Fn_Phi_2m1 = Sheets(DataSheetName).Range("ZZ1").Value

'      If Fn_Phi_2m1 < 0 Then
'         Fn_Phi_2m1 = 0
'      End If


End Function

'======================================================== <M1-1-F2>
'　関数：　代数螺旋の偏角φ2からφ1を求める
'
'
'========================================================

Public Function Fn_Phi_2m1_Solver(Phi_2 As Double, ByVal DataSheetName As String) As Double

    Dim I As Long, J As Long

    Sheets(DataSheetName).Activate

'-----------------
'　　Gaolseek φ2 → φ2m1 counter point between φ2 and φ1  (reverse wall contact angle)
'-----------------
'    With Sheets(DataSheetName)
'        .Range("ZZ1").Value = Phi_2 - pi                         ' phi2(i)  set Initial value to cell
'        .Range("ZY1").Value = Phi_2 - Atn(k / Phi_2) - pi        ' set tmporary value of θ to tenporary cell
'        .Range("ZY2").Value = k                                     ' Algebraic constat
'        .Range("AAA1").Formula = "=(ZZ1-atan(ZY2/ZZ1))-ZY1"         ' [Formura No.(12)]
'        .Range("AAA1").GoalSeek Goal:=0, ChangingCell:=Range("ZZ1") ' Goalseek
'
'    '      Range("AAA2").Value = phi2(i) - Atn(k / phi2(i)) - (the(i) + 2 * PI)  'φ2　検算表示
'    End With


'-----------------
'　　Solver φ2 → φ2m1 counter point between φ2 and φ1  (reverse wall contact angle)
'-----------------
      With Sheets(DataSheetName)
          .Range("ZZ1").Value = Phi_2 - pi                         ' phi2(i)  set Initial value to cell
          .Range("ZY1").Value = Phi_2 - Atn(k / Phi_2) - pi        ' set tmporary value of θ to tenporary cell
          .Range("ZY2").Value = k                                     ' Algebraic constat
          .Range("AAA1").Formula = "=(ZZ1-atan(ZY2/ZZ1))-ZY1"         ' [Formura No.(12)]
      End With

        Sheets(DataSheetName).Select
        SolverReset
'        SolverOptions MaxTime:=20, Precision:=1E-16

      '　制約条件  SolverAddで制約条件を設定。
      '     Relationは１が less than(≦)、２がequal＝、３が more than≧
        SolverAdd CellRef:="ZZ1", Relation:=3, FormulaText:="0.0000000000000001"
        ' SolverAdd CellRef:="ZZ1", Relation:=1, FormulaText:="PI()/2.5"

      ' Parameter設定
      '     MaxMinVal =2 （最小値にする=2；最大にする=1；特定値にする=3）
        SolverOk SetCell:="AAA1", MaxMinVal:=3, ValueOf:=0, ByChange:=Range("ZZ1"), _
            Engine:=1, EngineDesc:="GRG Nonlinear"

        SolverSolve UserFinish:=True    ' 結果ボックスを非表示

'-----------------
'   Goalseekの結果：φ2 を戻り値に設定
'-----------------
     Fn_Phi_2m1_Solver = Sheets(DataSheetName).Range("ZZ1").Value

'      If Fn_Phi_2m1 < 0 Then
'         Fn_Phi_2m1 = 0
'      End If


End Function

'======================================================== <M1-2>
'  Caculate Theta θ(rotation Angle) and Index ()　:
'
'
'    使用Sub　： Calc_Phi2_AB_Max_to_Min　← Goalseek 利用
'　  ・ 代数螺旋の偏角φ2､φ1を求める｡ また、設定角度幅毎のφ2とφ1を求める｡
'　  ・ 代数螺旋の偏角φ1、φ2に対する、軸角度θcを求める。
'
'========================================================

Public Sub Calc_Theta_Index_Number_set()            '　20171030

Dim I As Long, J As Long

'　Redim      　参考) ReDim Preserve

'--------------------------------------------
' [1] Wrap回転角 Max  the_C1 　θ ：　the
'
'　　・圧縮開始基準の軸回転角θc:the_c()と、φ2基準θ:のthe()は回転方向が逆
'     　Index No.は圧縮方向に増加する様に定義する為、
'     　圧縮方向に回転すると、Phi2()、the()は減少し、The_c()は増加する
'        　the_c()=the(0)-the()
'
'  　・代数螺旋偏角 P2_C1に対応する軸回転角The_cを基準に、
'　　　固定で刻み、配列 the(dw_n)に代入する。
'
        '    代数螺旋の偏角を固定で刻み、配列に代入
        '    　A室圧縮開始角度～AorB室最終側の圧縮終了角度
        '    偏角φ2   Phi_2_Amax >= Phi_2 　> (Phi_2_Amin or Phi_2_Bmin)
        '　　偏角φ2   PC_C1 　　 >  PC_C2 　> PC_C3 　> PC_C4
        '    Indexφ2　dw_n_C1 　 >  dw_n_C2 > dw_n_C3 > dw_n_C4
'--------------------------------------------
'
    dw_n_C1 = 0
    I = 0
    J = 0

    For I = 0 To dw_n
        the(I) = The_Max - dw * I

'                Debug.Print "[1] i/ theta = " & i & "/ " & the(i) * 180 / PI & "deg "
    Next I

'    Debug.Print "[C1n-0]  theta(0)    = " & the(0) * 180 / pi & " deg "
''    Debug.Print "　　 theta(" & dw_n_end & ") = " & the(dw_n_end) * 180 / pi & " deg "
'    Debug.Print "　　 theta(" & dw_n & ") = " & the(dw_n) * 180 / pi & " deg "
'    Debug.Print " "

''--------------------------------------------
'' [2] 代数螺旋偏角 the_C4  P2_C2
''      B室圧縮開始の偏角（φB_srt）の処理
''      φB_srtが、固定分割点と異なる場合は、φB_srtの分割点を加える
''--------------------------------------------
'    The_C2 = The_Cx :  dw_n_C2 = dw_n_Cx

''--------------------------------------------
'' [3] 代数螺旋偏角 the_C3  P2_C3
''--------------------------------------------
'    The_C3 = The_Cx   :   dw_n_C3 = dw_n_Cx

''--------------------------------------------
'' [4] 代数螺旋偏角 the_C4    P2_C4
''      最終圧縮時の偏角（φAB_end）の処理
''　　　φAB_endが、固定分割点と異なる場合は、φAB_endの分割点を加える
''--------------------------------------------
'    The_C4 = The_Cx :  dw_n_C4 = dw_n_Cx

'--------------------------------------------
' [1巻き目、２巻き目]
'   代数螺旋 偏角 the_C1n(),the_C2n(),the_C3n(),the_C4n()
'--------------------------------------------
      '[00] P2_C1_deg =18.3608637309803deg         A室 圧縮開始 偏角  P2
      '     P2_C2_deg =15.2192710773906deg         B室 圧縮開始 偏角  P2
      '     P2_C3_deg =8.51605781042989deg         A室 圧縮終了 偏角  P2 / P1=2.47836753783162
      '     P2_C4_deg =7.79220751400521deg         A室 圧縮終了 偏角  P2 / P1=1.850049007114

      '  The_C1n(0) = 18.311 / 1049.16263180024deg    [0]巻内側 A室 圧縮開始 回転角 /index=0
      '  The_C2n(0) = 15.160 / 868.578209507063deg    [0]巻内側 B室 圧縮開始 回転角 /index=91
      '  The_C3n(0) = 14.693 / 841.834862665542deg    [0]巻内側 A室 圧縮終了 回転角 /index=105
      '  The_C4n(0) = 13.959 / 799.799583703732deg    [0]巻内側 B室 圧縮終了 回転角 /index=127

      '  The_C1n(1) = 12.028 / 689.162631800239deg    [1]巻内側 A室 圧縮開始 回転角 /index=183 180?
      '  The_C2n(1) = 08.876 / 508.578209507063deg    [1]巻内側 B室 圧縮開始 回転角 /index=274 272?
      '  The_C3n(1) = 08.410 / 481.834862665542deg    [1]巻内側 A室 圧縮終了 回転角 /index=288 287?
      '  The_C4n(1) = 07.676 / 439.799583703732deg    [1]巻内側 B室 圧縮終了 回転角 /index=310

      '  The_C1n(2) = 05.745 / 329.162631800239deg    [2]巻内側 A室 圧縮開始 回転角 /index=366 360?
      '  The_C2n(2) = 02.593 / 148.578209507063deg    [2]巻内側 B室 圧縮開始 回転角 /index=0
      '  The_C3n(2) = 02.126 / 121.834862665542deg    [2]巻内側 A室 圧縮終了 回転角 /index=0
      '  The_C4n(2) = 01.393 / 79.7995837037325deg    [2]巻内側 B室 圧縮終了 回転角 /index=0

'--C1n
        The_C1n(0) = The_C1n(0)
        dw_n_C1n(0) = 0

        For I = 1 To 2
            If The_C1n(I) > the(dw_n) Then
                The_Cx = The_C1n(I)                   ' initial value初期値
                dw_n_Cx = 0                           ' initial value初期値
                  txt_Cx = "C1n-" & Format(I, "@")    ' for comment

                Call Calc_Theta_Index_Sorting

                The_C1n(I) = The_Cx
                dw_n_C1n(I) = dw_n_Cx
            End If
        Next I

'--C2n
        For I = 0 To 2
            If The_C2n(I) > the(dw_n) Then
                The_Cx = The_C2n(I)
                dw_n_Cx = 0
                txt_Cx = "C2n-" & Format(I, "@")

                Call Calc_Theta_Index_Sorting

                The_C2n(I) = The_Cx
                dw_n_C2n(I) = dw_n_Cx
            End If
        Next I

'--C3n
        For I = 0 To 2
            If The_C3n(I) > the(dw_n) Then
                The_Cx = The_C3n(I)
                dw_n_Cx = 0
                txt_Cx = "C3n-" & Format(I, "@")

                Call Calc_Theta_Index_Sorting

                The_C3n(I) = The_Cx
                dw_n_C3n(I) = dw_n_Cx
            End If
        Next I

'--C4n

        For I = 0 To 2
            If The_C4n(I) > the(dw_n) Then
                The_Cx = The_C4n(I)
                dw_n_Cx = 0
                txt_Cx = "C4n-" & Format(I, "@")

                Call Calc_Theta_Index_Sorting

                The_C4n(I) = The_Cx
                dw_n_C4n(I) = dw_n_Cx
            End If
        Next I


    For jj_a = 0 To dw_n

         ' θ ：the(0) 圧縮開始=φPhi基準の軸回転角度    [rad]
         ' θc：The_c()　圧縮開始基準 軸回転角度 0～     [rad]

            the_c(jj_a) = the(0) - the(jj_a)

        For I = 0 To 2
            If The_C1n(I) = the(jj_a) Then
               dw_n_C1n(I) = jj_a
            End If
            If The_C2n(I) = the(jj_a) Then
               dw_n_C2n(I) = jj_a
            End If
            If The_C3n(I) = the(jj_a) Then
               dw_n_C3n(I) = jj_a
            End If
            If The_C4n(I) = the(jj_a) Then
               dw_n_C4n(I) = jj_a
            End If
         Next I

    Next jj_a


      dw_n_C1 = dw_n_C1n(0)
      dw_n_C2 = dw_n_C2n(0)
      dw_n_C3 = dw_n_C3n(1)
      dw_n_C4 = dw_n_C4n(1)
      dw_n_end = dw_n_end - 1

'            Debug.Print "[" & txt_Cx & "] dw_n_" & txt_Cx & " /end  = " & dw_n_Cx & " / " & dw_n_end
''            Debug.Print "     dw_n_end     = " & dw_n_end
''            Debug.Print "     P2_Cx        = " & P2_Cx * 180 / PI & "deg "
'            Debug.Print "     the(dw_n_" & txt_Cx & ") = " & the(dw_n_Cx) * 180 / pi & "deg "
'            Debug.Print "     The_" & txt_Cx & "       = " & The_Cx * 180 / pi & "deg "



'---------------------------------------
'　　回転角Theta θ=0,π,2π,3π,4π,･･･ 半回転毎のIndex No.を dw_n_PI(j)に格納
'　　　　dw_n_PI(j) ⇒ jxπ回転、
'---------------------------------------

    ReDim dw_n_PI(turn_wrap_n * 2)

        For I = 0 To dw_n
            For J = 1 To turn_wrap_n * 2

                If (the(0) - the(I)) = pi * J Then
                    dw_n_PI(J) = I
                End If

            Next J
        Next I


End Sub


'======================================================== <M1-2>
'  Caculate Theta θ(rotation Angle) and Index ()　:
'
'
'    使用Sub　： Calc_Phi2_AB_Max_to_Min　← Goalseek 利用
'　  ・ 代数螺旋の偏角φ2､φ1を求める｡ また、設定角度幅毎のφ2とφ1を求める｡
'　  ・ 代数螺旋の偏角φ1、φ2に対する、軸角度θcを求める。
'
'========================================================

Public Sub Calc_Theta_Index_Number_set_check()            '　20180704

Dim I As Long, J As Long

'　Redim      　参考) ReDim Preserve

'--------------------------------------------
' [1] Wrap回転角 Max  the_C1 　θ ：　the
'
'　　・圧縮開始基準の軸回転角θc:the_c()と、φ2基準θ:のthe()は回転方向が逆
'     　Index No.は圧縮方向に増加する様に定義する為、
'     　圧縮方向に回転すると、Phi2()、the()は減少し、The_c()は増加する
'        　the_c()=the(0)-the()
'
'  　・代数螺旋偏角 P2_C1に対応する軸回転角The_cを基準に、
'　　　固定で刻み、配列 the(dw_n)に代入する。
'
        '    代数螺旋の偏角を固定で刻み、配列に代入
        '    　A室圧縮開始角度～AorB室最終側の圧縮終了角度
        '    偏角φ2   Phi_2_Amax >= Phi_2 　> (Phi_2_Amin or Phi_2_Bmin)
        '　　偏角φ2   PC_C1 　　 >  PC_C2 　> PC_C3 　> PC_C4
        '    Indexφ2　dw_n_C1 　 >  dw_n_C2 > dw_n_C3 > dw_n_C4
'--------------------------------------------
'
    dw_n_C1 = 0
    I = 0
    J = 0

'    Debug.Print "[C1n-0]  theta(0)    = " & the(0) * 180 / pi & " deg "
'    Debug.Print "　　     theta(" & dw_n & ") = " & the(dw_n) * 180 / pi & " deg "
'    Debug.Print " "


'--C1n
        The_C1n(0) = The_C1n(0)
        dw_n_C1n(0) = 0

        For I = 0 To 2
                The_Cx = The_C1n(I)                   ' initial value初期値
                dw_n_Cx = dw_n_C1n(I)                 ' initial value初期値
                txt_Cx = "C1n-" & Format(I, "@")      ' for comment

                'Call Calc_Theta_Index_Sorting_check
        Next I

'--C2n
         For I = 0 To 2
                The_Cx = The_C2n(I)                   ' initial value初期値
                dw_n_Cx = dw_n_C2n(I)                 ' initial value初期値
                txt_Cx = "C2n-" & Format(I, "@")      ' for comment

                'Call Calc_Theta_Index_Sorting_check
        Next I

'--C3n
        For I = 0 To 2
                The_Cx = The_C3n(I)                   ' initial value初期値
                dw_n_Cx = dw_n_C3n(I)                 ' initial value初期値
                txt_Cx = "C3n-" & Format(I, "@")      ' for comment

                'Call Calc_Theta_Index_Sorting_check
        Next I


'--C4n
         For I = 0 To 2
                The_Cx = The_C4n(I)                   ' initial value初期値
                dw_n_Cx = dw_n_C4n(I)                 ' initial value初期値
                txt_Cx = "C4n-" & Format(I, "@")      ' for comment

                'Call Calc_Theta_Index_Sorting_check
        Next I


End Sub


'======================================================== <M1-2-1>
'  軸回転角 the(i)  を降順に並べ替える
'
'　　the(0)＝0   at 圧縮開始(φ2= FS_in_Max、V_a_Max
'    the(end)= 2π*(巻数+1)
'
'========================================================

Public Sub Calc_Theta_Index_Sorting()            '　20171102

Dim I As Long, J As Long

'--------------------------------------------
' [x] 代数螺旋偏角 を降順に並べ替える
'         B室圧縮開始の偏角（φB_srt）の処理
'         φB_srtが、固定分割点と異なる場合は、φB_srtの分割点を加える
'--------------------------------------------

    dw_n_Cx = 0
    I = 0
    J = 0

    Do While (the(I) >= The_Cx) And (I <= dw_n)              '【PC_C2順番の抽出】
            dw_n_Cx = I
            I = I + 1
    Loop
'            Debug.Print "[" & txt_Cx & "] dw_n_" & txt_Cx & " /end  = " & dw_n_Cx & " / " & dw_n_end
''            Debug.Print "     dw_n_end     = " & dw_n_end
''            Debug.Print "     P2_Cx        = " & P2_Cx * 180 / PI & "deg "
'            Debug.Print "     the(dw_n_" & txt_Cx & ") = " & the(dw_n_Cx) * 180 / pi & "deg "
'            Debug.Print "     The_" & txt_Cx & "       = " & The_Cx * 180 / pi & "deg "

    If the(dw_n_Cx) = The_Cx Then

        dw_n_end = dw_n_end
        the(dw_n_Cx) = The_Cx

    ElseIf the(dw_n_Cx) > The_Cx Then

        dw_n_end = dw_n_end + 1
        dw_n_Cx = dw_n_Cx + 1

            For J = dw_n To dw_n_Cx + 1 Step -1          '【降順に並べ替え】
               the(J) = the(J - 1)
            Next J

        the(dw_n_Cx) = The_Cx

    Else
            Stop
    End If


'-- 表示　　　Index 入替番号と前後の3つを表示

'    Debug.Print "[" & txt_Cx & "]  dw_n_" & txt_Cx & " /end = " & dw_n_Cx & " / " & dw_n_end
'
'        For j = -1 To 1
'            Debug.Print "　　 Theta(" & dw_n_Cx + j & ")= " & the(dw_n_Cx + j) * 180 / pi & "deg "
'        Next j
'            Debug.Print " "

End Sub


'======================================================== <M1-2-1>
'  軸回転角 the(i)  を降順に並べ替える
'
'========================================================

Public Sub Calc_Theta_Index_Sorting_check()            '　20180704

Dim I As Long, J As Long


    I = 0
    J = 0
            Debug.Print "[" & txt_Cx & "] dw_n_" & txt_Cx & " /end  = " & dw_n_Cx & " / " & dw_n_end
            Debug.Print "     the(dw_n_" & txt_Cx & ") = " & the(dw_n_Cx) * 180 / pi & "deg "
            Debug.Print "     The_" & txt_Cx & "       = " & The_Cx * 180 / pi & "deg "

'-- 表示　　　Index 入替番号と前後の3つを表示

'    Debug.Print "[" & txt_Cx & "]  dw_n_" & txt_Cx & " /end = " & dw_n_Cx & " / " & dw_n_end

        For J = -1 To 1
            If (dw_n_Cx + J) >= 0 And (dw_n_Cx + J) <= dw_n Then

            Debug.Print "　　 Theta(" & dw_n_Cx + J & ")= " & the(dw_n_Cx + J) * 180 / pi & "deg "
            End If
        Next J
            Debug.Print " "

End Sub

'======================================================== 【M1-3】
'-- Cellから読込
'
'
'========================================================


Public Sub Calc_Theta_Index_Number_Read()

    Dim I As Long, J As Long

    '---------------------------------------
    '【Cellから読込み】
    '---------------------------------------

    For J = 0 To dw_n                            ' 20171107

        phi2(J) = Sheets(DataSheetName).Cells(4 + J, 24).Value            '【Cell X4からX370読込み】
        phi1(J) = Sheets(DataSheetName).Cells(4 + J, 26).Value            '【Cellから読込み】
'        phi2m1(j) = Sheets(DataSheetName).Cells(4 + j, 26).Value         '【Cellから読込み】20180319

      For I = 1 To 9
        Phi_c_fi(I, J) = Sheets(DataSheetName).Cells(4 + J, 23 + I).Value  '【Cell X4からAF370読込み】20180320

      Next I

    Next J

End Sub                                             ' END <M1-3>


'========================================================
'  Wrap Contact Number of Chamber A, B
'     N_wrap_a(j) : Number of A Wrap contact points at each index J
'     N_wrap_b(j) : Number of A Wrap contact points at each index J
'     N_wrap_max  : Max Number of Wrap contact points in range
'
'========================================================

Public Sub Calc_Theta_Index_to_Wrap_Number()

   Dim I As Long, J As Long
   Dim I1 As Long, J1 As Long

   For J = 0 To dw_n

   '0 - 2PI
     If (the_c(J) >= 0) And (the_c(J) < 2 * pi) Then

         ' Wrap Contact Number of Chamber A
           If Phi_c_fi(7, J) > Phi_1_Amin Then
                   N_wrap_a(J) = 4
               ElseIf Phi_c_fi(5, J) > Phi_1_Amin Then
                  N_wrap_a(J) = 3
               ElseIf Phi_c_fi(3, J) > Phi_1_Amin Then
                  N_wrap_a(J) = 2
               Else
                  Stop
           End If

        ' Wrap Contact Number of Chamber B
           If Phi_c_fi(7, J) > Phi_1_Bmin Then
                  N_wrap_b(J) = 4
                     If Phi_c_fi(1, J) > Phi_2_Bmax Then
                        N_wrap_b(J) = 3
                     End If
               ElseIf Phi_c_fi(5, J) > Phi_1_Bmin Then
                  N_wrap_b(J) = 3
                     If Phi_c_fi(1, J) > Phi_2_Bmax Then
                        N_wrap_b(J) = 2
                     End If
               ElseIf Phi_c_fi(3, J) > Phi_1_Bmin Then
                  N_wrap_b(J) = 2
                     If Phi_c_fi(1, J) > Phi_2_Bmax Then
                        N_wrap_b(J) = 1
                     End If
               Else
                  Stop
           End If

   ' 2PI-4PI
     ElseIf (the_c(J) >= 2 * pi) And (the_c(J) <= 4 * pi) Then

         ' Wrap Contact Number of Chamber A
           If Phi_c_fi(7, J) > Phi_1_Amin Then
                  N_wrap_a(J) = 4 + 1
               ElseIf Phi_c_fi(5, J) > Phi_1_Amin Then
                  N_wrap_a(J) = 3 + 1
               ElseIf Phi_c_fi(3, J) > Phi_1_Amin Then
                  N_wrap_a(J) = 2 + 1
               ElseIf Phi_c_fi(1, J) > Phi_1_Amin Then
                  N_wrap_a(J) = 1 + 1
               Else
                   Stop
           End If

        ' Wrap Contact Number of Chamber B
           If Phi_c_fi(7, J) > Phi_1_Bmin Then
                  N_wrap_b(J) = 4 + 1
                     If Phi_c_fi(1, J) > Phi_2_Bmax Then
                        N_wrap_b(J) = 3 + 1
                     End If
               ElseIf Phi_c_fi(5, J) > Phi_1_Bmin Then
                  N_wrap_b(J) = 3 + 1
                     If Phi_c_fi(1, J) > Phi_2_Bmax Then
                        N_wrap_b(J) = 2 + 1
                     End If
               ElseIf Phi_c_fi(3, J) > Phi_1_Bmin Then
                  N_wrap_b(J) = 2 + 1
                     If Phi_c_fi(1, J) > Phi_2_Bmax Then
                        N_wrap_b(J) = 1 + 1
                     End If
               ElseIf Phi_c_fi(1, J) > Phi_1_Bmin Then
                  N_wrap_b(J) = 1 + 1
                     If Phi_c_fi(1, J) > Phi_2_Bmax Then
                        N_wrap_b(J) = 0 + 1
                     End If
               Else
                   Stop
           End If

     Else
        Stop
     End If

   Next J



   ' Max
      N_wrap_max = N_wrap_a(0)      ' Max Number of Wrap contact point

      For J = 1 To dw_n
         If N_wrap_max < N_wrap_a(J) Then
            N_wrap_max = N_wrap_a(J)
         End If
         If N_wrap_max < N_wrap_b(J) Then
            N_wrap_max = N_wrap_b(J)
         End If
      Next J

GoTo lbel_Calc_Theta_Index_to_Wrap_Number_end


'--------------------------
'-- Data 一括貼付
'--------------------------

    Sheets(DataSheetName_2).Select
'    Sheets(DataSheetName_2).Cells.ClearContents                    '：Sheet全Cellの数式、文字列をクリア
    Sheets(DataSheetName_2).Range("AH1:AI999").ClearContents        '：指定Cellの数式、文字列をクリア

    I1 = 4         ' 貼付先の先頭セルの、行と列 Cells(i1, j1)
    J1 = 34         ' 参考) Cells(2, 1).Select = Range("A2").Select　　'Cells(3,8) = Range("H3")
        With Sheets(DataSheetName_2)
            .Range(Cells(I1, J1), Cells(I1 + dw_n, J1)).Value _
                = WorksheetFunction.Transpose(N_wrap_a)              '= N_wrap_a             '
            .Range(Cells(I1, J1 + 1), Cells(I1 + dw_n, J1 + 1)).Value _
                = WorksheetFunction.Transpose(N_wrap_b)              '= N_wrap_b
        End With

        With Sheets(DataSheetName_2)
            .Cells(1, J1).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

            .Cells(3, J1).Value = "N_wrap_a"
            .Cells(3, J1 + 1).Value = "N_wrap_b"
        End With
'
'         '1. Range("A1")→A1セル
'         '2. Range("A1:B3")→A1～B3セル範囲
'         '3. Range("A1,B3")→A1とB3セル
'         '6. Range(Cells(1, 1), Cells(3,2))→A1～B3セル範囲
'         '9. Range("名前定義")→名前定義のセル範囲
'         '10. Range(Rows(1), Rows(3)) →1～3行の範囲
'         '11. Range(Columns(1), Columns(3)) →1～3列の範囲
'

'■Label point
lbel_Calc_Theta_Index_to_Wrap_Number_end:

End Sub

'======================================================== 【M1-4】
'　回転角Theta θに対する偏角φ2、φ1を求める
'　回転角Theta θ｡｡｡｡ PI毎の接点Φを求める
'
'========================================================

Public Sub Calc_Theta_Index_to_Phi()

    Dim phi_tmp As Double
    Dim I As Long, J As Long
    Dim I1 As Long, J1 As Long
    Dim Imax As Long, Jmax As Long

'---------------------------------------
        ' Wrap_Start_angle_min(0) = Min(Wrap_Start_angle_min(1)-(4))
        ' Wrap_Start_angle_min(1) = 1.00660817056929E-02  'FS_in   Wrap_Start_angle_min FS in (g1) Phi_1
        ' Wrap_Start_angle_min(2) = 0.710188517448186     'FS_out  Wrap_Start_angle_min FS out (g1) Phi_2
        ' Wrap_Start_angle_min(3) = 1.00660817056929E-02  'OS_in   Wrap_Start_angle_min OS in (g2) Phi_1
        ' Wrap_Start_angle_min(4) = 0.710188517448186     'OS_out  Wrap_Start_angle_min OS out (g2) Phi_2

    For I = 0 To dw_n 'Step 50
            '255 To 259 'dw_n

         phi2(I) = Fn_Phi2_at_theta(the(I), DataSheetName)

'             phi1(i) = Fn_Phi_1(phi2(i), DataSheetName)
'             phi2m1(i) = Fn_Phi_2m1(phi2(i), DataSheetName)       '20180319

           If (pi <= the_c(I) And (the_c(I) <= (2 * pi))) Then
              Phi_c_fi(9, I) = Fn_Phi2_at_theta(the(I) + pi, DataSheetName)
           End If

         Phi_c_fi(1, I) = phi2(I)
         phi_tmp = phi2(I)
           J = 2

                Do While (phi_tmp >= OS_in_srt And phi_tmp >= 0)
                  Phi_c_fi(J, I) = Fn_Phi_2m1(phi_tmp, DataSheetName)
                  phi_tmp = Phi_c_fi(J, I)

                      '  If phi_tmp < 0 Then
                      '      phi_tmp = 0
                      '  End If

                  J = J + 1
                Loop

         phi1(I) = Phi_c_fi(3, I)

            Debug.Print "θ" & Format(I, "(000)");
            If I Mod 10 = 0 Then Debug.Print

    Next I

'            Debug.Print "I-"
'            Debug.Print "/" & Format(I, "000");
'            If I Mod 20 = 0 Then Debug.Print
'                        Debug.Print "/" & Format(I, "000");
'            If I Mod 10 = 0 Then Debug.Print
'------------------------
'  【Cellへ書出し＆計算】
'------------------------

        I1 = Range("W4").Row          ' tmp_cell = "W4" : I1 = Range(tmp_cell).Row
        J1 = Range("W4").Column
        Imax = UBound(Phi_c_fi, 2)    '= dw_n   Data array length of raws
        Jmax = UBound(Phi_c_fi, 1)    '= 9      Data array width of collums

'        DataSheetName = "DataSheet_2"
        With Sheets(DataSheetName)
          '--
            .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
                                    = WorksheetFunction.Transpose(Phi_c_fi)
          '-- time stap
            .Cells(2, J1 - 2).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")
            .Cells(3, J1 - 2).Value = "Index No."             '　"Index No."
            .Cells(3, J1 - 1).Value = "Theta_c[deg]"          '　"the_c"

            For J = 0 To Jmax
              .Cells(3, J1 + J).Value = "Phi_c_fi(" & J & ",i)"
            Next J

        End With

            Debug.Print "　"
            Debug.Print " <Calc_Theta_Index_to_Phi>" & Format(Time - srt_time_1, " HH:mm:ss")


End Sub

'======================================================== 【M1-4】
'　回転角Theta θに対する偏角φ2、φ1を求める
'　回転角Theta θ｡｡｡｡ PI毎の接点Φを求める
'
'========================================================

Public Sub Calc_Theta_Index_to_Phi_test()

Dim I As Long, J As Long, Jmax As Long
Dim phi_tmp As Double
Dim solver_use As Long

For I = 0 To n
    Phi_c_fi(0, I) = Fn_Phi_2m1(phi_tmp, DataSheetName)
    Stop
    Phi_c_fi(0, I) = Fn_Phi_2m1_Solver(phi_tmp, DataSheetName)
    Stop
Next I

Stop

End Sub




'======================================================== 【M1-4】
'　回転角Theta θに対する偏角φ2、φ1を求める
'　回転角Theta θ｡｡｡｡ PI毎の接点Φを求める
'
'========================================================

Public Sub Calc_Theta_Index_to_Phi_all_0()

Dim I As Long, J As Long, Jmax As Long
Dim phi_tmp As Double
Dim solver_use As Long

        ' Wrap_Start_angle_min(1) = 1.00660817056929E-02  'FS_in   Wrap_Start_angle_min FS in (g1) Phi_1
        ' Wrap_Start_angle_min(2) = 0.710188517448186     'FS_out  Wrap_Start_angle_min FS out (g1) Phi_2
        ' Wrap_Start_angle_min(3) = 1.00660817056929E-02  'OS_in   Wrap_Start_angle_min OS in (g2) Phi_1
        ' Wrap_Start_angle_min(4) = 0.710188517448186     'OS_out  Wrap_Start_angle_min OS out (g2) Phi_2

    solver_use = 0
    Jmax = 2

'    Jmax = 6
'        For I = 0 To dw_n
'               Phi_c_fi(Jmax, I) = 0
'               Phi_c_fi(Jmax + 1, I) = 0
'        Next I

    For I = 0 To dw_n

    ' phi2(I) = Fn_Phi2_at_theta(the(I), DataSheetName)
    ' phi1(i) = Fn_Phi_1(phi2(i), DataSheetName)
    ' phi2m1(i) = Fn_Phi_2m1(phi2(i), DataSheetName)       '20180319

    '      If (pi <= the_c(I) And (the_c(I) <= (2 * pi))) Then
    '         Phi_c_fi(9, I) = Fn_Phi2_at_theta(the(I) + pi, DataSheetName)
    '      End If

    ' 最外側：１番外側のFS,OS接点Phi_c_fi(1, I)＝φ2(I)
        ' Phi_c_fi(1, I) = phi2(I)
        ' phi_tmp = phi2(I)

        phi_tmp = Phi_c_fi(Jmax - 1, I)

    ' 外側から中央へ２番～J番以降FS,OS接点Phi_c_fi(J, I)＝φ2(I)
          J = Jmax

        Do While (phi_tmp >= Wrap_Start_angle_min(0))

            If solver_use = 0 Then
              Phi_c_fi(J, I) = Fn_Phi_2m1(phi_tmp, DataSheetName)
            Else
              Phi_c_fi(J, I) = Fn_Phi_2m1_Solver(phi_tmp, DataSheetName)
            End If

              phi_tmp = Phi_c_fi(J, I)
'                    If Phi_c_fi(J, I) < Wrap_Start_angle_min(0)) Then
'                       Phi_c_fi(J, I) = Wrap_Start_angle_min(0))
'                    End If
          J = J + 1
        Loop

         phi1(I) = Phi_c_fi(3, I)

         Debug.Print "θ(" & (Format(I, "000)"));
         If I Mod 10 = 0 Then Debug.Print

    Next I

End Sub

'======================================================== 【M1-4-F】
'　関数：回転角度 Theta θ から代数螺旋の偏角φ2を求める
'
'　　　　 Goalseek 利用
'========================================================

Public Function Fn_Phi2_at_theta(Theta_1 As Double, ByVal DataSheetName As String) As Double

    Dim I As Long, J As Long

    Sheets(DataSheetName).Activate

'-----------------
'　　Gaolseek theta θ→ φ2
'-----------------
    With Sheets(DataSheetName)
        .Range("ZZ1").Value = Phi_1 + 2 * pi                        ' phi2(i)  set Initial value to cell
        .Range("ZY1").Value = Theta_1               ' set tmporary value of θ to tenporary cell
        .Range("ZY2").Value = k                                     ' Algebraic constat
        .Range("AAA1").Formula = "=ZZ1-atan(ZY2/ZZ1)-ZY1"           ' [Formura No.(12)]
        .Range("AAA1").GoalSeek Goal:=0, ChangingCell:=Range("ZZ1") ' Goalseek

    End With

'-----------------
'   Goalseekの結果：φ2 を戻り値に設定
'-----------------
     Fn_Phi2_at_theta = Range("ZZ1").Value

End Function

'
''======================================================= 【M1-5】

'======================================================== 【M1-6】
'　A室：内側代数螺旋の偏角φ2,φ1 を配列に格納する
'
'   使用関数    Phi_1       = Fn_Phi_1(Phi_2, DataSheetName)
'   使用関数    V_a(jj_a)   = Fn_Calc_Volume_A(Phi_1, Phi_2)
'
'========================================================

Public Sub Calc_Phi2_Phi1()

'---------------------------------------
' A室容積計算　： 圧縮開始 Index jj_a＝0
'---------------------------------------
'    DataSheetName = "DataSheet_5"
    jj_a = 0

'---------------------------------------
' [A-0]  A,B 共通　代数螺旋 偏角φ1&φ2, β1&β2,　軸回転角θc
'
'  < 読書用 Flag1 >
'　　　　Flag_n1 = 1 "OFFで計算しCellへ書出し" ,
'　　　　　　　　 =0 "ONでCellから数値を読み込む"
'---------------------------------------

If Flag_n1 = 1 Then
     '------------------------
    '【圧縮終了角度以後の処理】：以後は終了時の値を維持  / 面積用θ,φ2,φ1  計算
    '------------------------
            Debug.Print "　sub◇ Calc_Phi2_Phi1() :(Flag_n1= 1)  " & Time & " "
            Debug.Print "[A]" & "0～" & dw_n

'                                                   'Index φ２偏角範囲、圧縮開始＝0～終了tmp_n1
    For jj_a = 0 To dw_n                            ' 20171031

       If jj_a > dw_n_C4 Then

            Phi_2 = phi2(dw_n_C4)
            Phi_1 = phi1(dw_n_C4)
            phi1(jj_a) = Phi_1
'       Else
'            Phi_2 = phi2(jj_a)
'            Phi_1 = Fn_Phi_1(Phi_2, DataSheetName)      ' Goalseek用のDatasheet名が必要
'            phi1(jj_a) = Phi_1

       End If
'            the(jj_a) = phi2(jj_a) - Atn(k / phi2(jj_a)) - 2 * PI

'            Debug.Print "　Phi2(" & jj_a & ")",         ' ■■ 表示

    Next jj_a

'    For jj_a = 0 To dw_n
'
'         ' θc：The_c()　圧縮開始基準 軸回転角度 0～
'         ' θ：the(0) 圧縮開始=φPhi基準の軸回転角度
'            the_c(jj_a) = the(0) - the(jj_a)
'
'    Next jj_a


'    '------------------------
'    '【Cellへ書出し＆計算】
'    '------------------------
'    For jj_a = 0 To dw_n
'
''        Debug.Print "[A-0] jj_a = " & jj_a & " / dw_n_C4=" & dw_n_C4
''        Debug.Print "[A]" & jj_a,
'
'        With Sheets(DataSheetName)                       '【Cellに書き出し】
'            .Cells(jj_a + 4, 22).Value = jj_a            '
'            .Cells(jj_a + 4, 23).Value = phi2(jj_a)    '
'            .Cells(jj_a + 4, 24).Value = phi1(jj_a)
'            .Cells(jj_a + 4, 25).Value = the_c(jj_a) / PI * 180  ' [deg]
'        End With
'
'    Next jj_a
'
'        With Sheets(DataSheetName)
'            .Cells(3, 22).Value = "Index No."             '　"Index No."
'            .Cells(3, 23).Value = "Phi2"                  '　"Phi2"
'            .Cells(3, 24).Value = "Phi1"                  '　"Phi1"
'            .Cells(3, 25).Value = "Theta_c[deg]"          '　"the_c"
'
'            .Cells(2, 22).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")
'
''            .Cells(2, 23).Value = Format(Now(), "yyyy/MM/DD")       '　"Date
''            .Cells(2, 23).Value = Format(Now(), "HH:mm:ss")         '　"Date
'
'        End With


'ElseIf Flag_n1 = 0 Then
'
'    '---------------------------------------
'    '【Cellから読込み】
'    '---------------------------------------
'
'        Debug.Print "　sub◇ Calc_Phi2_Phi1() :(Flag_n1=0 読込)  " & Time & " "
'        Debug.Print "[A]" & "0～" & dw_n & " data from Cell"
'
'    For jj_a = 0 To dw_n                            ' 20171031
'
'        Phi_1 = Sheets(DataSheetName).Cells(jj_a + 4, 24).Value         '【Cellから読込み】
'        phi1(jj_a) = Phi_1
''        the(jj_a) = Phi_1 - Atn(k / Phi_1)        '
''        the(jj_a) = phi2(jj_a) - Atn(k / phi2(jj_a)) - 2 * PI      '
'
''        Debug.Print "[A-φ]" & jj_a,
''            Debug.Print "[A-0] jj_a = " & jj_a & " / dw_n_C4=" & dw_n_C4
'
'    Next jj_a
'
'Else
'        Stop

End If

End Sub


'======================================================== 【M2-1】
'　A室：内側代数螺旋の偏角φ2(>φ1)から、容積を求める
'
'   使用関数    Phi_1       = Fn_Phi_1(Phi_2, DataSheetName)
'   使用関数    V_a(jj_a)   = Fn_Calc_Volume_A(Phi_1, Phi_2)
'   接点情報を配列に保管　　⇒　Get_xy_Chamber_A
'                               Get_xy_Contact_Point_2PI_A
'
'========================================================

Public Sub Calc_Phi_to_Volume_A()

'---------------------------------------
' A室容積計算　： 圧縮開始 Index jj_a＝0
'---------------------------------------

Dim I As Long, J As Long:
Dim tmp_n1 As Long
'Dim tmp_C1 As Double, tmp_C2 As Double, tmp_C3 As Double

' 作用力計算用
   I = N_wrap_max - 1 ' =2    '0:最外接点、1：１周内側、２：２周内側で中央側接点

    ReDim xfi_c(I, dw_n):  ReDim yfi_c(I, dw_n): ReDim RR_fi_c(I, dw_n)
    ReDim xmo_c(I, dw_n):  ReDim ymo_c(I, dw_n): ReDim RR_mo_c(I, dw_n)

    ReDim beta_fi_c(I, dw_n)
    ReDim del_fi_c(I, dw_n): ReDim del_mo_c(I, dw_n)
    ReDim psi_fi_c(I, dw_n): ReDim psi_mo_c(I, dw_n)
    ReDim V_a(dw_n)

' A : center of Chamber fiqure
    ReDim xg_a(I, dw_n):         ReDim yg_a(I, dw_n)              ' A Chanmber 図心


'    DataSheetName = "DataSheet_3"
    jj_a = 0

'---------------------------------------
' [A-0]  A室 圧縮区間 容積 ：代数螺旋　偏角φ2end 開始角 / φ2srt 終了角
'---------------------------------------
'
'      [暫定]　3区間に分けて計算
'              1) [圧縮開始～内側へ1巻目]～[2巻目～圧縮終了､temp_n1]　～４π
'                    -> 圧縮終了後、接触点は１巻外側へ移る
'              2) [内側1巻目～2巻目]～圧縮終了～４π
'                    → 1)区間の2π以降をコピー
'              3) [2巻目～圧縮終了]～４π
'                    → 1)区間の2π以降をコピー
'
'  <memo>
'           dw_n_C1 : 0 : Long
'           dw_n_C2 : 91 : Long
'           dw_n_C3 : 287 : Long
'           dw_n_C4 : 310 : Long
'
      '[00] P2_C1_deg =18.3608637309803deg         A室 圧縮開始 偏角  P2
      '     P2_C2_deg =15.2192710773906deg         B室 圧縮開始 偏角  P2
      '     P2_C3_deg =8.51605781042989deg         A室 圧縮終了 偏角  P2 / P1=2.47836753783162
      '     P2_C4_deg =7.79220751400521deg         A室 圧縮終了 偏角  P2 / P1=1.850049007114

      '  The_C1n(0) = 18.311 / 1049.16263180024deg    [0]巻内側 A室 圧縮開始 回転角 /index=0
      '  The_C2n(0) = 15.160 / 868.578209507063deg    [0]巻内側 B室 圧縮開始 回転角 /index=91
      '  The_C3n(0) = 14.693 / 841.834862665542deg    [0]巻内側 A室 圧縮終了 回転角 /index=105
      '  The_C4n(0) = 13.959 / 799.799583703732deg    [0]巻内側 B室 圧縮終了 回転角 /index=127

      '  The_C1n(1) = 12.028 / 689.162631800239deg    [1]巻内側 A室 圧縮開始 回転角 /index=183 180?
      '  The_C2n(1) = 08.876 / 508.578209507063deg    [1]巻内側 B室 圧縮開始 回転角 /index=274 272?
      '  The_C3n(1) = 08.410 / 481.834862665542deg    [1]巻内側 A室 圧縮終了 回転角 /index=288 287?
      '  The_C4n(1) = 07.676 / 439.799583703732deg    [1]巻内側 B室 圧縮終了 回転角 /index=310

      '  The_C1n(2) = 05.745 / 329.162631800239deg    [2]巻内側 A室 圧縮開始 回転角 /index=366 360?
      '  The_C2n(2) = 02.593 / 148.578209507063deg    [2]巻内側 B室 圧縮開始 回転角 /index=0
      '  The_C3n(2) = 02.126 / 121.834862665542deg    [2]巻内側 A室 圧縮終了 回転角 /index=0
      '  The_C4n(2) = 01.393 / 79.7995837037325deg    [2]巻内側 B室 圧縮終了 回転角 /index=0



'---------------------------------------
' [A-1]  A室 圧縮区間 容積 ：代数螺旋　偏角φ2end 開始角 / φ2srt 終了角
'        A室  圧縮開始から 0巻,1巻内側のA室接点情報を各配列に保管
'---------------------------------------
' 最小Index φ1　偏角配列 Index tmp_n1

'      If P2_C3 = Phi_2_Amin Then     'P2_C3_deg >= P2_C4_deg
'              tmp_n1 = dw_n_C3               ' dw_n_C3 = 287 ?
'          Else
'              tmp_n1 = dw_n_C4               ' dw_n_C4 = 310
'      End If

      tmp_n1 = dw_n_C3

  '[a-1] 圧縮区間での容積計算
     For jj_a = 0 To dw_n
'      For jj_a = 0 To tmp_n1          ' Index φ２偏角範囲、圧縮開始＝0～圧縮終了tmp_n1(=287)

             Phi_2 = phi2(jj_a)           'Phi_2 > Phi_1
             Phi_1 = phi1(jj_a)

         If Phi_1 > 0 Then
           '<M2-1-F1>
             V_a(jj_a) = Fn_Calc_Volume_A(Phi_1, Phi_2)

             xg_a(0, jj_a) = xg_a_tmp   '--- A Chamber: center of crescent area　(含む：offset dx,dy)
             yg_a(0, jj_a) = yg_a_tmp
   '          S_a(0, jj_a) = Area_tmp

             ' Debug.Print " V_a(" & jj_a & ")",                   '■ = " & V_a(jj_a)

            ' 圧縮区間 A室の接点情報を各配列に保管
             Call Get_xy_Chamber_A
                  '　圧縮室の巻終(Phi_2)、巻始(Phi_1)接点から、各0巻、1巻目の接点情報を保管
                  ' 接点座標x,y 、β2, ψfi2, δfi2,  RRfi2, / x,y ,β2, ψmo2, δmo2,  RRmo2
                  '      xfi_c(j, jj_a) = xfi(i) ,xmo_c(), beta_fi_c(), psi_fi_c(), del_fi_c()
                  '      yfi_c(j, jj_a) = yfi(i),
                  '      RR_fi_c(j, jj_a) = Sqr(RR_fi(i))

         Else        ': Stop
            '-- Formura (19) fi
               xfi_c(J, jj_a) = 0
               yfi_c(J, jj_a) = 0
               RR_fi_c(J, jj_a) = 0
            '-- Formura (21) mo
               xmo_c(J, jj_a) = 0
               ymo_c(J, jj_a) = 0
               RR_mo_c(J, jj_a) = 0

            '-- β1, θ, δfi1,Φfi1
               beta_fi_c(J, jj_a) = 0
            '-- FS_in  Formura (20)
               psi_fi_c(J, jj_a) = 0
               del_fi_c(J, jj_a) = 0
            '-- OS_out Formura (22)
               psi_mo_c(J, jj_a) = 0
               del_mo_c(J, jj_a) = 0

            '--- A Chamber: center of crescent area
               xg_a(J, jj_a) = 0
               yg_a(J, jj_a) = 0
         End If

      Next jj_a


'---------------------------------------
' [A-1b]  RR 動径処理
'---------------------------------------

      For jj_a = tmp_n1 + 1 To dw_n

            '-- Formura (19) fi
'               xfi_c(j, jj_a) = xfi(i)
'               yfi_c(j, jj_a) = yfi(i)
               RR_fi_c(0, jj_a) = RR_fi_c(0, jj_a - dw_n_PI(2))
            '-- Formura (21) mo
'               xmo_c(j, jj_a) = xmo(i)
'               ymo_c(j, jj_a) = ymo(i)
               RR_mo_c(0, jj_a) = RR_mo_c(0, jj_a - dw_n_PI(2))
            '-- Formura (19) fi
'               xfi_c(j, jj_a) = xfi(i)
'               yfi_c(j, jj_a) = yfi(i)
               RR_fi_c(1, jj_a) = RR_fi_c(1, jj_a - dw_n_PI(2))
            '-- Formura (21) mo
'               xmo_c(j, jj_a) = xmo(i)
'               ymo_c(j, jj_a) = ymo(i)
               RR_mo_c(1, jj_a) = RR_mo_c(1, jj_a - dw_n_PI(2))

      Next jj_a

'---------------------------------------
' [A-2]  A室  外側圧縮開始から　２巻内側～圧縮終了までの　A室接点情報を各配列に保管
'---------------------------------------
      If P2_C3 > 2 * pi Then
             Call Get_xy_Contact_Point_2PI_A
      Else
            Stop
      End If


'---------------------------------------
' [A-3]  圧縮終了後
'---------------------------------------

'      If P2_C3 < P2_C4 Then     'P2_C3 >= P2_C4
'              tmp_n1 = dw_n_C3               ' dw_n_C3 = 287 ?
'          Else
'              tmp_n1 = dw_n_C4               ' dw_n_C4 = 310
'      End If

   For jj_a = tmp_n1 + 1 To dw_n    ' Index φ２偏角範囲、圧縮終了tmp_n1(=287)～軸回転終了

        V_a(jj_a) = V_a(tmp_n1)

   Next jj_a

            V_a_Max = V_a(0)                '[mm3]
            V_a_Min = V_a(tmp_n1)           '[mm3]


        Debug.Print vbCrLf & "<M2-1>"
        Debug.Print " [A-1]A室 V_a_Max = " & V_a_Max / 1000 & "(" & 0 & ")"
        Debug.Print "          V_a_Min = " & V_a_Min / 1000 & "(" & tmp_n1 & ")"
         ' Debug.Print "          S_a_Max = " & V_a_Max / Hw & "S_a_Min = " & V_a_Min / Hw
        Debug.Print "    Volume Ratio A = " & V_a_Max / V_a_Min


End Sub


'======================================================== 【M2-1-F1】
'  関数：Caculate Chamber_A Volume
'
'========================================================

Public Function Fn_Calc_Volume_A(ByVal Phi_1 As Double, ByVal Phi_2 As Double) As Double

    Dim I As Long, J As Long
    Dim v1 As Long, v2 As Long
    Dim div_phi As Double

    Dim beta_tmp As Double:      Dim delta_tmp As Double

    Dim Sfi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Crescent_A As Double:                                   ' A Chamber crescent 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Crescent_A As Double:     Dim Sgy_Crescent_A As Double    ' A Chamber crescent 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

    Dim the_1 As Double:              Dim the_2 As Double             ' Phi_1 => the_1  Phi_2 => the_2

'------------------
'　A Chamber　　：配列設定
'------------------

    div_phi = (Phi_2 - Phi_1) / div_n    ' Divied angle　分割の角度幅

    ReDim phi1_v(div_n)
    ReDim psi_mo1_v(div_n):   ReDim psi_fi1_v(div_n)

    ReDim xfi(div_n):   ReDim yfi(div_n):  ReDim RR_fi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n):  ReDim RR_mo(div_n)

    ReDim beta_fi1_v(div_n)
    ReDim del_fi1_v(div_n)
    ReDim del_mo1_v(div_n)

'------------------
'　積分範囲 -- A Chamber
'------------------

For I = 0 To div_n

   '-- β1, θ, δfi1,Φfi1
        phi1_v(I) = Phi_1 + div_phi * I
           If phi1_v(I) = 0 Then
               phi1_v(I) = 0.0000000000001
               beta_tmp = Atn(1) * 2
           Else
               beta_tmp = Atn(k / phi1_v(I))      ' β1
           End If

        beta_fi1_v(I) = beta_tmp

     '-- FS_in  Formura (20)
        delta_tmp = -Atn(g1 * Sin(beta_tmp) / (a * phi1_v(I) ^ k + g1 * Cos(beta_tmp)))
        psi_fi1_v(I) = phi1_v(I) + delta_tmp
        del_fi1_v(I) = delta_tmp

     '-- OS_out Formura (22)
        delta_tmp = Atn(g2 * Sin(beta_tmp) / (a * phi1_v(I) ^ k - g2 * Cos(beta_tmp)))
        psi_mo1_v(I) = phi1_v(I) + delta_tmp
        del_mo1_v(I) = delta_tmp

   '------------------------------------------------
   '-- A Chamber
   '   包絡線の座標    ：[xfi(i) ,yfi(i)],[xmo(i) ,ymo(i)]
   '   包絡線の動径    ：RR_fi(i) , RR_mo(i)
   '------------------------------------------------
       ' Formura (8) fi
           xfi(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
           yfi(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy
           RR_fi(I) = xfi(I) ^ 2 + yfi(I) ^ 2

       ' Formura (5) mo
           xmo(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) - g2 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
           ymo(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) - g2 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy
           RR_mo(I) = xmo(I) ^ 2 + ymo(I) ^ 2

Next I

'------------------------------------------------
'-- A Chamber
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Smo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0
    the_1 = Phi_1 - Atn(k / Phi_1) - qq
    the_2 = Phi_2 - Atn(k / Phi_2) - qq

For I = 1 To div_n

   '---- A Chamber    '  Formura (13)
        del_Sfi = RR_fi(I) * (psi_fi1_v(I) - psi_fi1_v(I - 1))      '= 分割面積 x2 (注意
        Sfi_tmp = Sfi_tmp + del_Sfi                               '= 総和面積 x2 (注意

        del_Smo = RR_mo(I) * (psi_mo1_v(I) - psi_mo1_v(I - 1))
        Smo_tmp = Smo_tmp + del_Smo

   '----center of Chamber figure fi
        xg_fi_tmp = (xfi(I) + xfi(I - 1) + 0) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + 0) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi / 2) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi / 2) * xg_fi_tmp

   '----center of Chamber figure mo
        xg_mo_tmp = (xmo(I) + xmo(I - 1) + 0) / 3 + Ro * Cos(the_1)
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + 0) / 3 + Ro * Sin(the_1)
         ' xg_mo_tmp = (xmo(i) + xmo(i - 1) + 3 * Ro * Cos(the_1) + 0) / 3
         ' yg_mo_tmp = (ymo(i) + ymo(i - 1) + 3 * Ro * Sin(the_1) + 0) / 3
        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo / 2) * yg_mo_tmp    ' Sum Momet
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo / 2) * xg_mo_tmp

   '    Debug.Print "center of Chamber figure fi  i="; i

Next I


   '---- A Chamber Area   ' Formura (18) fi
        Crescent_A = (Sfi_tmp - Smo_tmp _
                + Sqr(RR_fi(0) * RR_fi(div_n)) * Sin(psi_fi1_v(0) + 2 * pi - psi_fi1_v(div_n)) _
                - Sqr(RR_mo(0) * RR_mo(div_n)) * Sin(psi_mo1_v(0) + 2 * pi - psi_mo1_v(div_n))) / 2

        Fn_Calc_Volume_A = Crescent_A * Hw              'A Chamber Volumr = 総和面積 * hight


   '--- Gravity center of Triangle at step area
        xg_fi_tmp = (xfi(0) + xfi(div_n) + 0) / 3
        yg_fi_tmp = (yfi(0) + yfi(div_n) + 0) / 3

        xg_mo_tmp = (xmo(0) + xmo(div_n) + 3 * Ro * Cos(the_1) + 0) / 3     '
        yg_mo_tmp = (ymo(0) + ymo(div_n) + 3 * Ro * Sin(the_1) + 0) / 3

   '--- A Chamber crescent area geometrical moment of area Sgx, Sgy 総和

    Sgx_Crescent_A = Sgx_fi_tmp - Sgx_mo_tmp _
      + Sqr(RR_fi(0) * RR_fi(div_n)) * Sin(psi_fi1_v(0) + 2 * pi - psi_fi1_v(div_n)) / 2 * yg_fi_tmp _
      - Sqr(RR_mo(0) * RR_mo(div_n)) * Sin(psi_mo1_v(0) + 2 * pi - psi_mo1_v(div_n)) / 2 * yg_mo_tmp

    Sgy_Crescent_A = Sgy_fi_tmp - Sgy_mo_tmp _
      + Sqr(RR_fi(0) * RR_fi(div_n)) * Sin(psi_fi1_v(0) + 2 * pi - psi_fi1_v(div_n)) / 2 * xg_fi_tmp _
      - Sqr(RR_mo(0) * RR_mo(div_n)) * Sin(psi_mo1_v(0) + 2 * pi - psi_mo1_v(div_n)) / 2 * xg_mo_tmp

   '--- A Chamber: center of crescent area

        xg_a_tmp = (Sgy_Crescent_A) / (Crescent_A) + dx
        yg_a_tmp = (Sgx_Crescent_A) / (Crescent_A) + dy
        Area_tmp = Crescent_A


End Function




'========================================================
' 配列に格納した接点を整理 巻数　： 圧縮開始 Index jj_a＝0
'
'========================================================


Public Sub Get_xy_Chamber_all()           '


' A室 0-１巻目 Index(0～dw_n_PI(2))　          / Phi2(0～2π) ：Get_xy_Chamber_A()
' 　　1-２巻目 Index(dw_n_PI(2)～dw_n_PI(4))   / Phi1(0～2π) ：Get_xy_Chamber_A()
' 　　2-３巻目 Index(dw_n_PI(4)～dw_n)       / Phi1(2π～min) ：Get_xy_Contact_Point_2PI_A()

Dim I As Long, J As Long:    Dim tmp_n1 As Long
Dim dw_n_Max As Long

'Dim tmp_C1 As Double, tmp_C2 As Double, tmp_C3 As Double

' 作用力計算用
   I = N_wrap_max - 1 ' =2    '0:最外接点、1：１周内側、２：２周内側で中央側接点
   dw_n_Max = dw_n + dw_n_PI(2)

    ReDim xfi_c(I, dw_n):  ReDim yfi_c(I, dw_n): ReDim RR_fi_c(I, dw_n)
    ReDim xmo_c(I, dw_n):  ReDim ymo_c(I, dw_n): ReDim RR_mo_c(I, dw_n)

    ReDim beta_fi_c(I, dw_n)
    ReDim del_fi_c(I, dw_n): ReDim del_mo_c(I, dw_n)
    ReDim psi_fi_c(I, dw_n): ReDim psi_mo_c(I, dw_n)
    ReDim V_a(dw_n)

' A : center of Chamber fiqure
    ReDim xg_a(I, dw_n):         ReDim yg_a(I, dw_n)              ' A Chanmber 図心


'---------------------------------------
' [A-1]  A室 圧縮区間
'
'
'---------------------------------------

  ' 最小Index φ1　偏角配列Index tmp_n1
'      If P2_C3_deg = Phi_2_Amin_deg Then     'P2_C3_deg >= P2_C4_deg
'              tmp_n1 = dw_n_C3
'          Else
'              tmp_n1 = dw_n_C4
'      End If

  ' 圧縮区間での容積計算
      I = 0
      For jj_a = 0 To dw_n_PI(2) - 1
         ' For jj_a = 0 To tmp_n1     ' Index φ２偏角範囲、圧縮開始＝0～圧縮終了tmp_n1(=287)
         ' jj_tmp = jj_a - dw_n_PI(2)         ' Index：2*PI内側

         Phi_2 = phi2(jj_a)       '  Phi_1 = phi1(jj_a)    'Phi_2 > Phi_1

           xfi_c(I, dw_n) = Fn_xfi(Phi_2)
           yfi_c(I, dw_n) = Fn_yfi(Phi_2)
           xmo_c(I, dw_n) = Fn_xmo(Phi_2)
           ymo_c(I, dw_n) = Fn_xmo(Phi_2)

           xfo_c(I, dw_n) = Fn_xfo(Phi_2)
           yfo_c(I, dw_n) = Fn_xfo(Phi_2)
           xmi_c(I, dw_n) = Fn_xmi(Phi_2)
           ymi_c(I, dw_n) = Fn_xmi(Phi_2)

      Next jj_a


      For jj_a = 0 To dw_n   ' For jj_a = 0 To tmp_n1     ' Index φ２偏角範囲、圧縮開始＝0～圧縮終了tmp_n1(=287)

         Phi_1 = phi1(jj_a)           'Phi_2 > Phi_1

         If Phi_1 > 0 Then

         I = 0
           xfi_c(I, dw_n + dw_n_PI(2)) = Fn_xfi(Phi_1)
           yfi_c(I, dw_n + dw_n_PI(2)) = Fn_yfi(Phi_1)
           xmo_c(I, dw_n + dw_n_PI(2)) = Fn_xmo(Phi_1)
           ymo_c(I, dw_n + dw_n_PI(2)) = Fn_xmo(Phi_1)

           xfo_c(I, dw_n + dw_n_PI(2)) = Fn_xfo(Phi_1)
           yfo_c(I, dw_n + dw_n_PI(2)) = Fn_xfo(Phi_1)
           xmi_c(I, dw_n + dw_n_PI(2)) = Fn_xmi(Phi_1)
           ymi_c(I, dw_n + dw_n_PI(2)) = Fn_xmi(Phi_1)

         I = 1
           xfi_c(I, dw_n) = Fn_xfi(Phi_1)
           yfi_c(I, dw_n) = Fn_yfi(Phi_1)
           xmo_c(I, dw_n) = Fn_xmo(Phi_1)
           ymo_c(I, dw_n) = Fn_xmo(Phi_1)

           xfo_c(I, dw_n) = Fn_xfo(Phi_1)
           yfo_c(I, dw_n) = Fn_xfo(Phi_1)
           xmi_c(I, dw_n) = Fn_xmi(Phi_1)
           ymi_c(I, dw_n) = Fn_xmi(Phi_1)

         End If

      Next jj_a

         If Phi_1 > 0 Then
           '<M2-1-F1>
             V_a(jj_a) = Fn_Calc_Volume_A(Phi_1, Phi_2)

             xg_a(0, jj_a) = xg_a_tmp   '--- A Chamber: center of crescent area
             yg_a(0, jj_a) = yg_a_tmp
   '          S_a(0, jj_a) = Area_tmp

             Debug.Print " V_a(" & jj_a & ")",                   '■ = " & V_a(jj_a)

            ' 圧縮区間 A室の接点情報を各配列に保管
             Call Get_xy_Chamber_A
                  ' 接点座標x,y 、β2, ψfi2, δfi2,  RRfi2, / x,y ,β2, ψmo2, δmo2,  RRmo2
                  '      xfi_c(j, jj_a) = xfi(i) ,xmo_c(), beta_fi_c(), psi_fi_c(), del_fi_c()
                  '      yfi_c(j, jj_a) = yfi(i),
                  '      RR_fi_c(j, jj_a) = Sqr(RR_fi(i)),

         Else: Stop
         End If




'---------------------------------------
' [A-3]  A室 吐出後　接点無区間
'---------------------------------------

   For jj_a = tmp_n1 + 1 To dw_n    ' Index φ２偏角範囲、圧縮終了tmp_n1(=287)～軸回転終了
      V_a(jj_a) = V_a(tmp_n1)

      '  吐出後区間 A室の接点情報を配列に保管
      Call Get_xy_Chamber_A

   Next jj_a

            V_a_Max = V_a(0)                '[mm3]
            V_a_Min = V_a(tmp_n1)           '[mm3]


        Debug.Print vbCrLf & "<M2-1>"
        Debug.Print " [A-1]A室 V_a_Max = " & V_a_Max / 1000 & "(" & 0 & ")"
        Debug.Print "          V_a_Min = " & V_a_Min / 1000 & "(" & tmp_n1 & ")"
         ' Debug.Print "          S_a_Max = " & V_a_Max / Hw & "S_a_Min = " & V_a_Min / Hw
        Debug.Print "    Volume Ratio A = " & V_a_Max / V_a_Min



End Sub

'========================================================
' 容積室の中央側接点座標、他　： 圧縮開始 Index jj_a＝0
' A室 2-３巻目 Index(dw_n_PI(4)～dw_n)/ Phi1(2π～min) ：Get_xy_Contact_Point_2PI_A()
   '++++++++++++++++++++++++++++++++++++++++
   '  A容積室の中央接点
   '  接点座標、β1, ψfi1, δfi1,  RRfi1, x,y
   '            β1, ψmo1, δmo1,  RRmo1, x,y
   '++++++++++++++++++++++++++++++++++++++++
'========================================================

Public Sub Get_xy_Contact_Point_2PI_A()

Dim I As Long, jj_tmp As Long, J As Long

' 2巻き目配列処理　 圧縮開始から１回転区間の接点　： 1巻目配列から配列コピー

   For jj_a = 0 To dw_n_PI(2)                ' Index of φPhi_1 偏角から2*PI内側

         J = 2    ' 2巻き目配列 Phi1　Index(dw_n_PI(4)～dw_n)/ Phi1(2π～4π)

         jj_tmp = jj_a + dw_n_PI(2)         ' Index：2*PI内側

         If jj_tmp <= dw_n Then

            '-- β1, θ, δfi1,Φfi1
               beta_fi_c(J, jj_a) = beta_fi_c(J - 1, jj_tmp)
            '-- FS_in  Formura (20)
               psi_fi_c(J, jj_a) = psi_fi_c(J - 1, jj_tmp)
               del_fi_c(J, jj_a) = del_fi_c(J - 1, jj_tmp)
            '-- OS_out Formura (22)
               psi_mo_c(J, jj_a) = psi_mo_c(J - 1, jj_tmp)
               del_mo_c(J, jj_a) = del_mo_c(J - 1, jj_tmp)

            '-- Formura (19) fi
               xfi_c(J, jj_a) = xfi_c(J - 1, jj_tmp)
               yfi_c(J, jj_a) = yfi_c(J - 1, jj_tmp)
               RR_fi_c(J, jj_a) = RR_fi_c(J - 1, jj_tmp)
            '-- Formura (21) mo
               xmo_c(J, jj_a) = xmo_c(J - 1, jj_tmp)
               ymo_c(J, jj_a) = ymo_c(J - 1, jj_tmp)
               RR_mo_c(J, jj_a) = RR_mo_c(J - 1, jj_tmp)

            '--- A Chamber: center of crescent area
               xg_a(J, jj_a) = xg_a(J - 1, jj_tmp)
               yg_a(J, jj_a) = yg_a(J - 1, jj_tmp)

         End If

   Next jj_a

End Sub


'========================================================
' 容積室の中央側接点座標、他　： 圧縮開始 Index jj_a＝0
' B室 2-３巻目 Index(dw_n_PI(4)～dw_n)/ Phi1(2π～min) ：Get_xy_Contact_Point_2PI_B()
   '++++++++++++++++++++++++++++++++++++++++
   '  B容積室の中央接点
   '  接点座標、β1, ψfi1, δfi1,  RRfi1, x,y
   '            β1, ψmo1, δmo1,  RRmo1, x,y
   '++++++++++++++++++++++++++++++++++++++++
'========================================================

Public Sub Get_xy_Contact_Point_2PI_B()


Dim I As Long, J As Long, jj_tmp As Long

   For jj_b = 0 To dw_n_PI(2)

         J = 2    ' Phi1
         jj_tmp = jj_b + dw_n_PI(2)

         If jj_tmp <= dw_n Then

            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = beta_mi_c(J - 1, jj_tmp)
            '-- OS_in  Formura (15)
               psi_mi_c(J, jj_b) = psi_mi_c(J - 1, jj_tmp)
               del_mi_c(J, jj_b) = del_mi_c(J - 1, jj_tmp)
            '-- FS_out Formura (17)
               psi_fo_c(J, jj_b) = psi_fo_c(J - 1, jj_tmp)
               del_fo_c(J, jj_b) = del_fo_c(J - 1, jj_tmp)

            '-- Formura (14) mi
               xmi_c(J, jj_b) = xmi_c(J - 1, jj_tmp)
               ymi_c(J, jj_b) = ymi_c(J - 1, jj_tmp)
               RR_mi_c(J, jj_b) = RR_mi_c(J - 1, jj_tmp)
            '-- Formura (16) fo
               xfo_c(J, jj_b) = xfo_c(J - 1, jj_tmp)
               yfo_c(J, jj_b) = yfo_c(J - 1, jj_tmp)
               RR_fo_c(J, jj_b) = RR_fo_c(J - 1, jj_tmp)

            '--- B Chamber: center of crescent area
               xg_b(J, jj_b) = xg_b(J - 1, jj_tmp)
               yg_b(J, jj_b) = xg_b(J - 1, jj_tmp)

         Else
            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = 0
            '-- OS_in  Formura (15)
               psi_mi_c(J, jj_b) = 0
               del_mi_c(J, jj_b) = 0
            '-- FS_out Formura (17)
               psi_fo_c(J, jj_b) = 0
               del_fo_c(J, jj_b) = 0

            '-- Formura (14) mi
               xmi_c(J, jj_b) = 0
               ymi_c(J, jj_b) = 0
               RR_mi_c(J, jj_b) = 0
            '-- Formura (16) fo
               xfo_c(J, jj_b) = 0
               yfo_c(J, jj_b) = 0
               RR_fo_c(J, jj_b) = 0

            '--- B Chamber: center of crescent area
               xg_b(J, jj_a) = 0
               yg_b(J, jj_a) = 0

         End If

   Next jj_b

End Sub


'========================================================
' 容積室の中央側接点座標、他　： 圧縮開始 Index jj_a＝0
' B室 0-0.5巻目 Index(0～dw_n_PI(1))　        / Phi1(0～π)　Get_xy_Contact_Point_B0()
'
'========================================================

Public Sub Get_xy_Contact_Point_B0()
   '++++++++++++++++++++++++++++++++++++++++ 圧縮開始からの容積室
   '  B容積室の中央接点
   '  接点座標、β1, ψfi1, δfi1,  RRfi1, x,y
   '            β1, ψmo1, δmo1,  RRmo1, x,y
   '++++++++++++++++++++++++++++++++++++++++

Dim J As Long

   For jj_b = 0 To dw_n_C2 - 1
         J = 0    ' Phi1

            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = beta_mi_c(J + 1, jj_b)
            '-- OS_in  Formura (15)
               psi_mi_c(J, jj_b) = psi_mi_c(J + 1, jj_b)
               del_mi_c(J, jj_b) = del_mi_c(J + 1, jj_b)
            '-- FS_out Formura (17)
               psi_fo_c(J, jj_b) = psi_fo_c(J + 1, jj_b)
               del_fo_c(J, jj_b) = del_fo_c(J + 1, jj_b)

            '-- Formura (14) mi
               xmi_c(J, jj_b) = xmi_c(J + 1, jj_b)
               ymi_c(J, jj_b) = ymi_c(J + 1, jj_b)
               RR_mi_c(J, jj_b) = RR_mi_c(J + 1, jj_b)
            '-- Formura (16) fo
               xfo_c(J, jj_b) = xfo_c(J + 1, jj_b)
               yfo_c(J, jj_b) = yfo_c(J + 1, jj_b)
               RR_fo_c(J, jj_b) = RR_fo_c(J + 1, jj_b)

   Next jj_b

End Sub


Public Sub Get_xy_Contact_Point_B1()
   '++++++++++++++++++++++++++++++++++++++++ 1巻内側の容積室
   '  B-1室  (A室先行圧縮区間）配列の入換え
   '  接点座標、β1, ψfi1, δfi1,  RRfi1, x,y
   '            β1, ψmo1, δmo1,  RRmo1, x,y
   '++++++++++++++++++++++++++++++++++++++++

Dim I As Long, J As Long, jj_tmp As Long

   For jj_b = 0 To dw_n_C2 - 1

         J = 1    ' Phi1
         jj_tmp = jj_b + dw_n_PI(2)

         If jj_tmp <= dw_n Then

            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = beta_mi_c(J, jj_tmp)
            '-- OS_in  Formura (15)
               psi_mi_c(J, jj_b) = psi_mi_c(J, jj_tmp)
               del_mi_c(J, jj_b) = del_mi_c(J, jj_tmp)
            '-- FS_out Formura (17)
               psi_fo_c(J, jj_b) = psi_fo_c(J, jj_tmp)
               del_fo_c(J, jj_b) = del_fo_c(J, jj_tmp)

            '-- Formura (14) mi
               xmi_c(J, jj_b) = xmi_c(J, jj_tmp)
               ymi_c(J, jj_b) = ymi_c(J, jj_tmp)
               RR_mi_c(J, jj_b) = RR_mi_c(J, jj_tmp)
            '-- Formura (16) fo
               xfo_c(J, jj_b) = xfo_c(J, jj_tmp)
               yfo_c(J, jj_b) = yfo_c(J, jj_tmp)
               RR_fo_c(J, jj_b) = RR_fo_c(J, jj_tmp)

         End If

   Next jj_b


End Sub


Public Sub Get_xy_Contact_Point_B2()
   '++++++++++++++++++++++++++++++++++++++++ 2巻内側の容積室
   '  B-2室  (A室先行圧縮区間）配列の入換え
   '  接点座標、β1, ψfi1, δfi1,  RRfi1, x,y
   '            β1, ψmo1, δmo1,  RRmo1, x,y
   '++++++++++++++++++++++++++++++++++++++++

Dim I As Long, J As Long, jj_tmp As Long

   For jj_b = 0 To dw_n_C2 - 1

         J = 2    ' Phi1
         jj_tmp = jj_b + dw_n_PI(2)

         If jj_tmp <= dw_n Then

            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = beta_mi_c(J, jj_tmp)
            '-- OS_in  Formura (15)
               psi_mi_c(J, jj_b) = psi_mi_c(J, jj_tmp)
               del_mi_c(J, jj_b) = del_mi_c(J, jj_tmp)
            '-- FS_out Formura (17)
               psi_fo_c(J, jj_b) = psi_fo_c(J, jj_tmp)
               del_fo_c(J, jj_b) = del_fo_c(J, jj_tmp)

            '-- Formura (14) mi
               xmi_c(J, jj_b) = xmi_c(J, jj_tmp)
               ymi_c(J, jj_b) = ymi_c(J, jj_tmp)
               RR_mi_c(J, jj_b) = RR_mi_c(J, jj_tmp)
            '-- Formura (16) fo
               xfo_c(J, jj_b) = xfo_c(J, jj_tmp)
               yfo_c(J, jj_b) = yfo_c(J, jj_tmp)
               RR_fo_c(J, jj_b) = RR_fo_c(J, jj_tmp)

         End If

   Next jj_b


End Sub

'======================================================== 【M2-2】
'　Ｂ室：内側代数螺旋の偏角φ2(>φ1)から、　B室容積を求める
'
'    使用関数   Phi_1       = Fn_Phi_1(Phi_2, DataSheetName)
'    使用関数   V_b(jj_b)   = Fn_Calc_Volume_B(Phi_1, Phi_2)
'   接点情報を配列に保管　　⇒　Get_xy_Chamber_B
'                               Get_xy_Contact_Point_2PI_B
'
'========================================================

Public Sub Calc_Phi_to_Volume_B()

'---------------------------------------
' B ： 変数設定    ：角度－Ａ室圧縮開始基準
'---------------------------------------

Dim I As Long, J As Long:    Dim tmp_n1 As Long

' 作用力計算用
   I = 2    '0:最外接点、1：１周内側、２：２周内側で中央側接点
    ReDim xmi_c(I, dw_n):  ReDim ymi_c(I, dw_n): ReDim RR_mi_c(I, dw_n)
    ReDim xfo_c(I, dw_n):  ReDim yfo_c(I, dw_n): ReDim RR_fo_c(I, dw_n)

    ReDim beta_mi_c(I, dw_n)
    ReDim del_mi_c(I, dw_n): ReDim del_fo_c(I, dw_n)
    ReDim psi_mi_c(I, dw_n): ReDim psi_fo_c(I, dw_n)
    ReDim V_b(dw_n)

' B : center of Chamber fiqure
    ReDim xg_b(I, dw_n):         ReDim yg_b(I, dw_n)              ' B Chanmber 図心

    jj_b = 0


'---------------------------------------
' [B-0]  B室 容積　圧縮区間　代数螺旋偏角φ2開始角
'                  　　　　　代数螺旋偏角φ2終了角
'---------------------------------------

  ' 最小Index φ1　偏角配列Index tmp_n1

'        tmp_n1 = Phi_2_Bmin

'   If Phi_1_Amin = Phi_1_Bmin Then     'If P2_C3 = P2_C4 Then
'           tmp_n1 = dw_n_C3
'      ElseIf P2_C4 = Phi_2_Bmin Then
'           tmp_n1 = dw_n_C4
'      Else
'           tmp_n1 = dw_n_C3
'   End If

      tmp_n1 = dw_n_C4

  ' 圧縮区間での容積計算
     For jj_b = 0 To dw_n
  '   For jj_b = 0 To tmp_n1

             Phi_2 = phi2(jj_b)
             Phi_1 = phi1(jj_b)

         If Phi_1 > 0 Then
           '<M2-2-F1>
             V_b(jj_b) = Fn_Calc_Volume_B(Phi_1, Phi_2)

                xg_b(0, jj_b) = xg_b_tmp   '--- B Chamber: center of crescent area　(含む：offset dx,dy)
                yg_b(0, jj_b) = yg_b_tmp

                ' Debug.Print " V_b(" & jj_b & ") ",       ' =" & V_b(jj_b)

           '接点情報代入 + (A室先行区間)B室の各接点情報＝０代入
            Call Get_xy_Chamber_B

         Else
            '-- Formura (14) mi
               xmi_c(J, jj_b) = 0
               ymi_c(J, jj_b) = 0
               RR_mi_c(J, jj_b) = 0
            '-- Formura (7) fo
               xfo_c(J, jj_b) = 0
               yfo_c(J, jj_b) = 0
               RR_fo_c(J, jj_b) = 0

            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = 0
            '-- OS_in  Formura (15)
               psi_mi_c(J, jj_b) = 0
               del_mi_c(J, jj_b) = 0
            '-- FS_out Formura (17)
               psi_fo_c(J, jj_b) = 0
               del_fo_c(J, jj_b) = 0

            '--- B Chamber: center of crescent area
               xg_b(J, jj_b) = 0
               yg_b(J, jj_b) = 0

         End If

   Next jj_b

'---------------------------------------
' [B-1]  B室 容積 (A室先行圧縮区間）
'             = 容積一定= B室吸込容積に置換
'---------------------------------------

   Phi_2 = phi2(dw_n_C2)
   Phi_1 = phi1(dw_n_C2)
   V_b(dw_n_C2) = Fn_Calc_Volume_B(Phi_1, Phi_2)

   For jj_b = 0 To dw_n_C2 - 1

       V_b(jj_b) = V_b(dw_n_C2)

   Next jj_b


 '---------------------------------------
' [B-1b]  RR 動径処理
'---------------------------------------

      For jj_b = tmp_n1 + 1 To dw_n

            '-- Formura (14) mi
            '   xmi_c(j, jj_b) = xmi(i)
            '   ymi_c(j, jj_b) = ymi(i)
               RR_mi_c(0, jj_b) = RR_mi_c(0, jj_b - dw_n_PI(2))
            '-- Formura (7) fo
            '   xfo_c(j, jj_b) = xfo(i)
            '   yfo_c(j, jj_b) = yfo(i)
               RR_fo_c(0, jj_b) = RR_fo_c(0, jj_b - dw_n_PI(2))
            '-- Formura (14) mi
            '   xmi_c(j, jj_b) = xmi(i)
            '   ymi_c(j, jj_b) = ymi(i)
               RR_mi_c(1, jj_b) = RR_mi_c(1, jj_b - dw_n_PI(2))
            '-- Formura (7) fo
            '   xfo_c(j, jj_b) = xfo(i)
            '   yfo_c(j, jj_b) = yfo(i)
               RR_fo_c(1, jj_b) = RR_fo_c(1, jj_b - dw_n_PI(2))
      Next jj_b


'---------------------------------------
' [B-2]  B室  圧縮開始から１回転区間　２巻内側のA室接点情報を各配列に保管
'
'---------------------------------------

   Call Get_xy_Contact_Point_2PI_B


'---------------------------------------
' [B-3]  B室  (A室先行圧縮区間）配列の入換え
'
'---------------------------------------
'
'   Call Get_xy_Contact_Point_B0  '-- B室(A室先行圧縮区間）配列の入換え
'
'   Call Get_xy_Contact_Point_B1  '-- B室(A室先行圧縮区間）配列の入換え
'
'   Call Get_xy_Contact_Point_B2  '-- B室(A室先行圧縮区間）配列の入換え

'---------------------------------------
' [B-4]  B室 吐出室連通区間
'            = 容積一定= B室最小容積に置換
'---------------------------------------

   Phi_2 = phi2(tmp_n1)
   Phi_1 = phi1(tmp_n1)

   V_b(tmp_n1) = Fn_Calc_Volume_B(Phi_1, Phi_2)

   For jj_b = tmp_n1 + 1 To dw_n       ' 角度、B 圧縮開始　＝0

       V_b(jj_b) = V_b(tmp_n1)

   Next jj_b

            V_b_Max = V_b(dw_n_C2)   '[mm3]
            V_b_Min = V_b(tmp_n1)    '[mm3]

        Debug.Print vbCrLf & "<M2-2>"
        Debug.Print " [B-1]B室  V_b_Max = " & V_b_Max / 1000 & "(" & dw_n_C2 & ")"
        Debug.Print "           V_b_Min = " & V_b_Min / 1000 & "(" & tmp_n1 & ")"
'        Debug.Print "          S_b_Max = " & V_b_Max / Hw & "S_b_Min = " & V_b_Min / Hw
        Debug.Print "    Volume Ratio B = " & V_b_Max / V_b_Min
        Debug.Print "    　　Volume A+B = " & (V_a_Max + V_b_Max) / 1000 & "cc"


End Sub



'======================================================== 【M2-2-F1】
'  関数：Fn_Calc_Volume_B()
'        Caculate Chamber_B Volume
'
'
'========================================================
'
'　積分範囲　   ：Integrate Range　　        ： Ψ1～Ψ2(=Φ1～Φ2)　　包絡線の偏角
'　分割数　     ：Division number　　　　    ： div_n
'　刻み幅(角度) ：calculation pitch width　  ： div_angle
'
'　　[1]　求める容積室の軸回転角度θ(又は、代数螺旋の偏角φ1,φ2)から、包絡線の偏角Ψ1Ψ2を決める
'　　[2]　代数螺旋の偏角φ1,φ2から、包絡線の動径R1､R2を求める
'    [3]　動径と刻み幅(角度)を用い台形積分し面積を求める。
'
'------------------------
'  Caculate Chamber_B Volume
'------------------------

Public Function Fn_Calc_Volume_B(ByVal Phi_1 As Double, ByVal Phi_2 As Double) As Double

    Dim I As Long, J As Long
    Dim v1 As Long, v2 As Long
    Dim div_phi As Double

    Dim beta_tmp As Double:        Dim delta_tmp As Double

    Dim Smi_tmp As Double:         Dim Sfo_tmp As Double
    Dim del_Smi As Double:          Dim del_Sfo As Double
    Dim Crescent_B As Double:

' B : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Crescent_B As Double:     Dim Sgy_Crescent_B As Double    ' A Chamber crescent 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

    Dim the_1 As Double:              Dim the_2 As Double             ' Phi_1 => the_1  Phi_2 => the_2

'------------------
'　B Chamber　　：配列設定
'         div_n : Chanber面積の分割数　台形公式利用
'------------------

    div_phi = (Phi_2 - Phi_1) / div_n               ' Divied angle　分割の角度幅

    ReDim phi1_v(div_n):
    ReDim psi_mi1_v(div_n):     ReDim psi_fo1_v(div_n)

    ReDim xmi(div_n):  ReDim ymi(div_n):  ReDim RR_mi(div_n)
    ReDim xfo(div_n):  ReDim yfo(div_n):  ReDim RR_fo(div_n)

    ReDim beta_mi1_v(div_n)
    ReDim del_mi1_v(div_n)
    ReDim del_fo1_v(div_n)

'------------------
'　積分範囲 -- B Chamber
'------------------

For I = 0 To div_n

'-- β1, θ, δmi1,Φmi1
     phi1_v(I) = Phi_1 + div_phi * I
        If phi1_v(I) = 0 Then
            phi1_v(I) = 0.0000000000001
            beta_tmp = Atn(1) * 2
        Else
            beta_tmp = Atn(k / phi1_v(I))      ' β1
        End If

     beta_mi1_v(I) = beta_tmp

  '-- OS_in  Formura (15)
     delta_tmp = -Atn(g2 * Sin(beta_tmp) / (a * phi1_v(I) ^ k + g2 * Cos(beta_tmp)))
     psi_mi1_v(I) = phi1_v(I) + pi + delta_tmp
     del_mi1_v(I) = delta_tmp

  '-- FS_out  Formura (17)
     delta_tmp = Atn(g1 * Sin(beta_tmp) / (a * phi1_v(I) ^ k - g1 * Cos(beta_tmp)))
     psi_fo1_v(I) = phi1_v(I) + pi + delta_tmp
     del_fo1_v(I) = delta_tmp

'------------------------------------------------
'-- B Chamber
'         包絡線の座標    ：[xmi(i) ,ymi(i)], [xfo(i) ,yfo(i)]
'         包絡線の動径    ：RR_mi(i) , RR_fo(i)
'------------------------------------------------
    ' Formura (6) mi
        xmi(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) - g2 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        ymi(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) - g2 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy
        RR_mi(I) = xmi(I) ^ 2 + ymi(I) ^ 2

   ' Formura (7) fo
        xfo(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        yfo(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy
        RR_fo(I) = xfo(I) ^ 2 + yfo(I) ^ 2

Next I

'------------------------------------------------
'-- B Chamber
'    各動径間の分割面積総和  ：del_Smi() , del_Sfo()
'------------------------------------------------

    Smi_tmp = 0
    Sfo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0
    the_1 = Phi_1 - Atn(k / Phi_1) - qq
    the_2 = Phi_2 - Atn(k / Phi_2) - qq

For I = 1 To div_n

 '---- B Chamber      ' Formura (13)
        del_Smi = RR_mi(I) * (psi_mi1_v(I) - psi_mi1_v(I - 1))      '= 分割面積 x2 (注意
        Smi_tmp = Smi_tmp + del_Smi                               '= 総和面積 x2 (注意

        del_Sfo = RR_fo(I) * (psi_fo1_v(I) - psi_fo1_v(I - 1))
        Sfo_tmp = Sfo_tmp + del_Sfo

 '----center of Chamber fiqure mi
        xg_mi_tmp = (xmi(I) + xmi(I - 1) + 3 * Ro * Cos(the_1) + 0) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + 3 * Ro * Sin(the_1) + 0) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + del_Smi / 2 * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + del_Smi / 2 * xg_mi_tmp

'----center of Chamber fiqure fo
        xg_fo_tmp = (xfo(I) + xfo(I - 1) + 0) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + 0) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + del_Sfo / 2 * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + del_Sfo / 2 * xg_fo_tmp

 Next I


'---- B Chamber    ' Formura (13) mi
        Crescent_B = (Smi_tmp - Sfo_tmp _
                + Sqr(RR_mi(0) * RR_mi(div_n)) * Sin(psi_mi1_v(0) + 2 * pi - psi_mi1_v(div_n)) _
                - Sqr(RR_fo(0) * RR_fo(div_n)) * Sin(psi_fo1_v(0) + 2 * pi - psi_fo1_v(div_n))) / 2

        Fn_Calc_Volume_B = Crescent_B * Hw

'--- Gravity center of Triangle at step area
        xg_mi_tmp = (xmi(0) + xmi(div_n) + 3 * Ro * Cos(the_1) + 0) / 3
        yg_mi_tmp = (ymi(0) + ymi(div_n) + 3 * Ro * Sin(the_1) + 0) / 3

        xg_fo_tmp = (xfo(0) + xfo(div_n) + 0) / 3
        yg_fo_tmp = (yfo(0) + yfo(div_n) + 0) / 3

'--- B Chamber crescent area geometrical moment of area Sgx, Sgy

    Sgx_Crescent_B = Sgx_mi_tmp - Sgx_fo_tmp _
      + Sqr(RR_mi(0) * RR_mi(div_n)) * Sin(psi_mi1_v(0) + 2 * pi - psi_mi1_v(div_n)) / 2 * yg_mi_tmp _
      - Sqr(RR_fo(0) * RR_fo(div_n)) * Sin(psi_fo1_v(0) + 2 * pi - psi_fo1_v(div_n)) / 2 * yg_fo_tmp

    Sgy_Crescent_B = Sgy_mi_tmp - Sgy_fo_tmp _
      + Sqr(RR_mi(0) * RR_mi(div_n)) * Sin(psi_mi1_v(0) + 2 * pi - psi_mi1_v(div_n)) / 2 * xg_mi_tmp _
      - Sqr(RR_fo(0) * RR_fo(div_n)) * Sin(psi_fo1_v(0) + 2 * pi - psi_fo1_v(div_n)) / 2 * xg_fo_tmp

 '--- B Chamber: center of crescent area

        xg_b_tmp = (Sgy_Crescent_B) / (Crescent_B) + dx
        yg_b_tmp = (Sgx_Crescent_B) / (Crescent_B) + dy
        Area_tmp = Crescent_B

End Function

'========================================================
' A室 0-１巻目 Index(0～dw_n_PI(2))　          は、Phi2(0～2π) ：Get_xy_Chamber_A()
' 　　1-２巻目 Index(dw_n_PI(2)～dw_n_PI(4))   は、Phi1(0～2π) ：Get_xy_Chamber_A()
'========================================================

Public Sub Get_xy_Chamber_A()
    Dim I As Long, J As Long

      '---------------------------------------
      '  容積室の外周側接点    0巻目
      '  接点座標、β2, ψfi2, δfi2,  RRfi2, x,y
      '            β2, ψmo2, δmo2,  RRmo2, x,y
      '---------------------------------------
         J = 0       ' 　 　0=外周側接点　1:中央側接点,
         I = div_n   ' Phi2側接点　　※面積計算時の上限、終点

            '-- Formura (19) fi
               xfi_c(J, jj_a) = xfi(I)
               yfi_c(J, jj_a) = yfi(I)
               RR_fi_c(J, jj_a) = Sqr(RR_fi(I))
            '-- Formura (21) mo
               xmo_c(J, jj_a) = xmo(I)
               ymo_c(J, jj_a) = ymo(I)
               RR_mo_c(J, jj_a) = Sqr(RR_mo(I))

            '-- β1, θ, δfi1,Φfi1
               beta_fi_c(J, jj_a) = beta_fi1_v(I)
            '-- FS_in  Formura (20)
               psi_fi_c(J, jj_a) = psi_fi1_v(I)
               del_fi_c(J, jj_a) = del_fi1_v(I)
            '-- OS_out Formura (22)
               psi_mo_c(J, jj_a) = psi_mo1_v(I)
               del_mo_c(J, jj_a) = del_mo1_v(I)

            '--- A Chamber: center of crescent area
               xg_a(J, jj_a) = xg_a_tmp
               yg_a(J, jj_a) = yg_a_tmp

      '---------------------------------------
      '  容積室の中央側接点    1巻目
      '  接点座標、β1, ψfi1, δfi1,  RRfi1, x,y
      '            β1, ψmo1, δmo1,  RRmo1, x,y
      '---------------------------------------
         J = 1    ' 0=外周側接点　1:中央側接点,
         I = 0    ' Phi1側接点   ※面積計算時の下限、始点

            '-- Formura (19) fi
               xfi_c(J, jj_a) = xfi(I)
               yfi_c(J, jj_a) = yfi(I)
               RR_fi_c(J, jj_a) = Sqr(RR_fi(I))
            '-- Formura (21) mo
               xmo_c(J, jj_a) = xmo(I)
               ymo_c(J, jj_a) = ymo(I)
               RR_mo_c(J, jj_a) = Sqr(RR_mo(I))

            '-- β1, θ, δfi1,Φfi1
               beta_fi_c(J, jj_a) = beta_fi1_v(I)
            '-- FS_in  Formura (20)
               psi_fi_c(J, jj_a) = psi_fi1_v(I)
               del_fi_c(J, jj_a) = del_fi1_v(I)
            '-- OS_out Formura (22)
               psi_mo_c(J, jj_a) = psi_mo1_v(I)
               del_mo_c(J, jj_a) = del_mo1_v(I)

            '--- A Chamber: center of crescent area
               xg_a(J, jj_a) = xg_a_tmp
               yg_a(J, jj_a) = yg_a_tmp


End Sub

'========================================================
' B室の接点
' B室 0.5-1巻目 Index(dw_n_PI(1)～dw_n_PI(2)) / Phi2(π～2π) ：Get_xy_Chamber_B()
' 　　1.5-2巻目 Index(dw_n_PI(2)～dw_n_PI(4)) / Phi1(0～2π)　：Get_xy_Chamber_B()
'========================================================

Public Sub Get_xy_Chamber_B()

    Dim I As Long, J As Long

      '---------------------------------------
      '  B容積室の外周側接点
      '  接点座標、β2, ψfi2, δfi2,  RRfi2, x,y
      '            β2, ψmo2, δmo2,  RRmo2, x,y
      '---------------------------------------
         J = 0       ' 　 　0=外周側接点　1:中央側接点,
         I = div_n   ' Phi2側接点　　※面積計算時の上限、終点

         If jj_b >= dw_n_C2 Then

            '-- Formura (14) mi
               xmi_c(J, jj_b) = xmi(I)
               ymi_c(J, jj_b) = ymi(I)
               RR_mi_c(J, jj_b) = Sqr(RR_mi(I))
            '-- Formura (7) fo
               xfo_c(J, jj_b) = xfo(I)
               yfo_c(J, jj_b) = yfo(I)
               RR_fo_c(J, jj_b) = Sqr(RR_fo(I))

            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = beta_mi1_v(I)
            '-- OS_in  Formura (15)
               psi_mi_c(J, jj_b) = psi_mi1_v(I)
               del_mi_c(J, jj_b) = del_mi1_v(I)
            '-- FS_out Formura (17)
               psi_fo_c(J, jj_b) = psi_fo1_v(I)
               del_fo_c(J, jj_b) = del_fo1_v(I)

            '--- B Chamber: center of crescent area
               xg_b(J, jj_b) = xg_b_tmp
               yg_b(J, jj_b) = yg_b_tmp

         Else                       ' B室=0、圧縮開始までの処理
            '-- Formura (14) mi
               xmi_c(J, jj_b) = 0
               ymi_c(J, jj_b) = 0
               RR_mi_c(J, jj_b) = 0
            '-- Formura (7) fo
               xfo_c(J, jj_b) = 0
               yfo_c(J, jj_b) = 0
               RR_fo_c(J, jj_b) = 0

            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = 0
            '-- OS_in  Formura (15)
               psi_mi_c(J, jj_b) = 0
               del_mi_c(J, jj_b) = 0
            '-- FS_out Formura (17)
               psi_fo_c(J, jj_b) = 0
               del_fo_c(J, jj_b) = 0

            '--- B Chamber: center of crescent area
               xg_b(J, jj_b) = 0
               yg_b(J, jj_b) = 0

         End If

      '---------------------------------------
      '  B容積室の中央側接点  吐出区間
      '  接点座標、β1, ψfi1, δfi1,  RRfi1, x,y
      '            β1, ψmo1, δmo1,  RRmo1, x,y
      '---------------------------------------
         J = 1    ' 0=外周側接点　1:中央側接点,
         I = 0    ' Phi1側接点   ※面積計算時の下限、始点

            '-- Formura (14) fi
               xmi_c(J, jj_b) = xmi(I)
               ymi_c(J, jj_b) = ymi(I)
               RR_mi_c(J, jj_b) = Sqr(RR_mi(I))
            '-- Formura (16) mo
               xfo_c(J, jj_b) = xfo(I)
               yfo_c(J, jj_b) = yfo(I)
               RR_fo_c(J, jj_b) = Sqr(RR_fo(I))

            '-- β1, θ, δmi1,Φmi1
               beta_mi_c(J, jj_b) = beta_mi1_v(I)
            '-- FS_in  Formura (15)
               psi_mi_c(J, jj_b) = psi_mi1_v(I)
               del_mi_c(J, jj_b) = del_mi1_v(I)
            '-- OS_out Formura (17)
               psi_fo_c(J, jj_b) = psi_fo1_v(I)
               del_fo_c(J, jj_b) = del_fo1_v(I)

            '--- B Chamber: center of crescent area
               xg_b(J, jj_b) = xg_b_tmp
               yg_b(J, jj_b) = yg_b_tmp


End Sub

'======================================================== 【M2-4】
'　Dicharge室：内側代数螺旋の偏角φ2(>φ1)から、吐出弁前室の容積と面積及び重心を求める
'
'   使用関数    Phi_1       = Fn_Phi_1(Phi_2, DataSheetName)
'   使用関数    V_d(jj_a)   = Fn_Calc_Volume_D(Phi_1, Phi_2)
'   接点情報を配列に保管　　⇒　Get_xy_Chamber_D
'                               Get_xy_Contact_Point_2PI_D
'
'========================================================

Public Sub Calc_Phi_to_Volume_D()

    Dim I As Long, J As Long

' A : center of Chamber fiqure
'    ReDim xg_d(dw_n):         ReDim yg_d(dw_n)              ' A Chanmber 図心
'    ReDim Sg_d(dw_n):
'
'    Dim the_1 As Double:

'-----------
 For I = 0 To 0   'dw_n

'   the_1 = the(i) - qq

'-----------
'[tg2 FS]
       Phi_2 = Phi_c_fi(4, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(5, I)
       Phi_0 = Phi_c_fi(6, I)

     '<A side FS_in>   Half moon
          Call Get_Gravity_Center_Chamber_D_FS(Phi_0, Phi_1)

          xg_f(0, I) = xg_a_tmp
          yg_f(0, I) = yg_a_tmp
          Sg_f(0, I) = Area_tmp

    ' <B side OS_in>  half moon
          Call Get_Gravity_Center_Chamber_D_OS(Phi_0, Phi_1)

          xg_m(0, I) = xg_a_tmp + Ro * Cos(the_1)
          yg_m(0, I) = yg_a_tmp + Ro * Sin(the_1)
          Sg_m(0, I) = Area_tmp


''[tg1 FS]  from Calc_Gravity_Center_wrap
'          xg_f(1, i) = xg_a_tmp
'          yg_f(1, i) = yg_a_tmp
'          Sg_f(1, i) = Area_tmp
'
''[tp1 OS]  from Calc_Gravity_Center_wrap
'          xg_m(1, i) = xg_a_tmp + Ro * Cos(the_1)
'          yg_m(1, i) = yg_a_tmp + Ro * Sin(the_1)
'          Sg_m(1, i) = Area_tmp


 '[ppi-Pg(1)] Discharge chamber

          Sg2_a(0, I) = Sg_f(0, I) + Sg_m(0, I) - Sg_f(1, I) - Sg_m(1, I)

          xg2_a(0, I) = (Sg_f(0, I) * xg_f(0, I) + Sg_m(0, I) * xg_m(0, I) _
                        - Sg_f(1, I) * xg_f(1, I) - Sg_m(1, I) * xg_m(1, I)) / Sg2_a(0, I)

          yg2_a(0, I) = (Sg_f(0, I) * yg_f(0, I) + Sg_m(0, I) * yg_m(0, I) _
                        - Sg_f(1, I) * yg_f(1, I) - Sg_m(1, I) * yg_m(1, I)) / Sg2_a(0, I)

  Next

End Sub



'======================================================== 【M3-1】
'
'   圧縮室A,B 圧力
'     Chamber No. 0->1,2,3    外(suction)->中央(discharge)
'
'      Press_A(0, i) = P_suction
'      Press_B(0, i) = P_suction
'
'========================================================

Sub Calc_Press_AB()

Dim I As Long, J As Long
    Dim tmp_vp1 As Double:          Dim tmp_vp2 As Double
    Dim Index_v_Amin As Long:       Dim Index_v_Bmin As Long

'   j = 3
      J = N_wrap_max                     ' Max Number of Wrap contact point
    ReDim Press_A(J, dw_n):          ReDim Press_B(J, dw_n)

'    ReDim Press_A(dw_n, n_ft):   ReDim Press_B(dw_n, n_ft)
'    turn_wrap_n = Application.RoundUp(FS_in_end / (2 * PI), 0)

   For I = 0 To dw_n
      Press_A(0, I) = P_suction
      Press_B(0, I) = P_suction
   Next I

    If FS_in_srt > OS_in_srt Then
        Index_v_Amin = dw_n_C3
        Index_v_Bmin = dw_n_C4
    Else
        Index_v_Bmin = dw_n_C3
        Index_v_Amin = dw_n_C4
    End If


'-----------------
' A_0 室 圧縮開始前後の圧力処理
'   Va <= Va_min then Pa = Ps
'   Va >= Va_max then Pa = Pd
'-----------------

 J = 1
    For I = 0 To dw_n

        If V_a(I) <= 0 Then
            Press_A(J, I) = Press_A(0, 0)

        ElseIf (V_a_Min <= V_a(I)) And (V_a(I) <= V_a_Max) And (I <= Index_v_Amin) Then

            Press_A(J, I) = Press_A(0, 0) * (V_a_Max / V_a(I)) ^ kappa

            If Press_A(J, I) > P_discharge Then
                Press_A(J, I) = P_discharge
            End If
        Else
            Press_A(J, I) = P_discharge
        End If

    Next I


'-----------------
' B_0 室 圧縮開始前後の圧力処理
'   Vb <= Vb_min then Pa = Ps
'   Vb >= Vb_max then Pa = Pd
'-----------------

    For I = 0 To dw_n

        If V_b(I) <= 0 Then
            Press_B(J, I) = Press_B(0, 0)

        ElseIf (V_b_Min <= V_b(I)) And (V_b(I) <= V_b_Max) And (I <= Index_v_Bmin) Then

            Press_B(J, I) = Press_B(0, 0) * (V_b_Max / V_b(I)) ^ kappa

            If Press_B(J, I) > P_discharge Then
                  Press_B(J, I) = P_discharge
            End If
        Else
            Press_B(J, I) = P_discharge
        End If

    Next I

'-----------------
'　＜ 1巻き内側＞
' A_1 室 圧縮開始前後の圧力処理
'   Va <= Va_min then Pa = Ps
'   Va >= Va_max then Pa = Pd
' B_1 室 圧縮開始前後の圧力処理
'   Vb <= Vb_min then Pa = Ps
'   Vb >= Vb_max then Pa = Pd
'-----------------

   '--
    J = 2

    For I = dw_n_PI(2) To dw_n
            Press_A(J, I - dw_n_PI(2)) = Press_A(J - 1, I)
            Press_B(J, I - dw_n_PI(2)) = Press_B(J - 1, I)
    Next I

    For I = 0 To dw_n
        If Press_A(J, I) = 0 Then
            Press_A(J, I) = P_discharge
        End If
        If Press_B(J, I) = 0 Then
            Press_B(J, I) = P_discharge
        End If

    Next I


   '--
    J = 3

    For I = 0 To dw_n

        If N_wrap_a(I) >= 1 And N_wrap_a(I) < 4 Then
            Press_A(J, I) = P_discharge
        ElseIf N_wrap_a(I) >= 4 Then
            Stop
        End If

        If N_wrap_b(I) >= 1 And N_wrap_b(I) < 4 Then
            Press_B(J, I) = P_discharge
        ElseIf N_wrap_b(I) >= 4 Then
            Stop
        End If

    Next I




End Sub



'======================================================== 【M5-1】
'　Wrap_thickness
'
'
'========================================================

Public Sub Calc_Wrap_thickness()

'--------------------------
' FS Wrap 厚み計算　： 圧縮開始　jj_a＝0
'--------------------------

Dim I As Long, J As Long
Dim tmp_n1 As Long

'Dim tmp_C0 As Double
Dim tmp_C1 As Double, tmp_C2 As Variant
Dim tmp_C3 As Double, tmp_C4 As Double

'    DataSheetName = "DataSheet_4"
    jj_a = 0
'   dw_t       ' [rad]　Wrap厚み計算の分割幅 dw_t_deg=15deg

'--------------------------
' -- 配列長を決める
'--------------------------

        tmp_n1 = (OS_in_end - (30 * pi / 180)) / dw_t
'        tmp_n1 = (OS_in_end_deg - OS_in_srt_deg) / tmp_C0
        tmp_n1 = WorksheetFunction.RoundUp(tmp_n1, 0)

        ReDim phi_iw(tmp_n1):           ReDim phi_ow(tmp_n1)
        ReDim gzai_fs(tmp_n1):          ReDim gzai_os(tmp_n1)
        ReDim thick_fs(tmp_n1):         ReDim thick_os(tmp_n1)

        ReDim w_xfi(tmp_n1):            ReDim w_yfi(tmp_n1)
        ReDim w_xfo(tmp_n1):            ReDim w_yfo(tmp_n1)
        ReDim w_xmi(tmp_n1):            ReDim w_ymi(tmp_n1)
        ReDim w_xmo(tmp_n1):            ReDim w_ymo(tmp_n1)

'    Debug.Print "■ Wrap厚 φout 計算Start => time= " & Time
'    Debug.Print " i=0 to " & tmp_n1

    For I = 0 To tmp_n1
      phi_iw(I) = OS_in_end - (dw_t) * I

    Next I



If Flag_n1 = 1 Then

    '--------------------------
      'Solver の式と制約条件を設定する。
       ' SolverReset
       ' SolverOptions
       ' SolverOk:       目標条件を設定する｡
       ' SolverSolve:       ソルバーを実行する｡
       ' SolverFinish: 終了処理。求めた解を該当セル(B1欄)に書き込む。
    '--------------------------

        phi_i = phi_iw(0)
        Phi_o = phi_i + pi * 0.9999    ' 初期値

    With Sheets(DataSheetName)
        .Range("ZZ1").Value = Phi_o                         ' Phi_o : φ2out Initial value
        .Range("ZY1").Value = phi_i                         ' Phi_i : φ2in
        .Range("ZY2").Value = k                             ' Algebraic constat
        .Range("ZY3").Value = a                             ' Algebraic constat

        .Range("AAA2").Formula = "=(ZY3*(ZY1^ZY2*COS(ZY1)+ZZ1^ZY2*COS(ZZ1)))/(-COS(ZY1-ATAN(ZY2/ZY1))+COS(ZZ1-ATAN(ZY2/ZZ1)))"   ' ' ξ’の分子/分母
        .Range("AAA4").Formula = "=ZY3*(ZY1^ZY2*sin(ZY1)+ZZ1^ZY2*sin(ZZ1))"         ' a * (Phi_i^(k) * Sin(Phi_i) + Phi_o^(k) * Sin(Phi_o))
        .Range("AAA5").Formula = "=sin(ZY1-atan(ZY2/ZY1))-sin(ZZ1-atan(ZY2/ZZ1))"   ' Sin(Phi_i-atan(k/Phi_i)) - Sin(Phi_o-atan(k/Phi_o))
        .Range("AAA1").Formula = "=(AAA4+(AAA2)*AAA5)^2"
    End With

        Sheets(DataSheetName).Select
        SolverReset

    '--------------------------
    '　制約条件  SolverAddで制約条件を設定。Relationは１が≦、２が＝、３が≧
    '--------------------------

    '        SolverOptions MaxTime:=0, Iterations:=200, Precision:=0.0000000000001, _
    '            Convergence:=0.00001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, _
    '            Derivatives:=2

        Sheets(DataSheetName).Select
        SolverAdd CellRef:="$ZZ$1", Relation:=1, FormulaText:="$ZY$1+3.14*0.999999"
        SolverAdd CellRef:="$ZZ$1", Relation:=3, FormulaText:="$ZY$1+3.14*0.6"
        SolverAdd CellRef:="$AAA$2", Relation:=1, FormulaText:="12"
        SolverAdd CellRef:="$AAA$2", Relation:=3, FormulaText:="0.001"

    '--------------------------
    '　各偏角φiで、φoの収束解をSolverで求める
    '--------------------------
    '
  For jj_a = 0 To tmp_n1      'dw_n_C4          ' 回転軸角度、圧縮開始　＝0
        '    For jj_a = 0 To 10        'tmp_n1      'dw_n_C4          ' 回転軸角度、圧縮開始　＝0

        phi_i = phi_iw(jj_a)
        Phi_o = phi_i + pi * 0.999            ' initial value

        With Sheets(DataSheetName)
            .Range("ZY1").Value = phi_i                         ' Phi_i : φ2in
            .Range("ZZ1").Value = Phi_o                         ' Phi_o : φ2out Initial value

        End With

    '--------------------------
    '   Solver
    '--------------------------
        Sheets(DataSheetName).Select

        Call Calc_Solve_thickness

            tmp_C1 = (Phi_o - phi_i) / pi * 180                     ' 差
            tmp_C2 = Sheets(DataSheetName).Range("AAA1").Value      ' 収束誤差
            tmp_C3 = Sheets(DataSheetName).Range("ZZ1").Value
            tmp_C4 = Sheets(DataSheetName).Range("AAA2").Value

    '--------------------------
    '--  収束エラー時の処理
    '--------------------------

    If tmp_C2 > 0.0001 Or tmp_C3 <= 0 Or tmp_C2 > 0.1 Then       ' 収束エラー時の処理
            With Sheets(DataSheetName)
                .Range("ZY1").Value = phi_i                        ' Phi_i : φ2in
                .Range("ZZ1").Value = phi_i + pi * 0.88             ' Phi_o : φ2out Initial value
            End With
        Call Calc_Solve_thickness

    End If

    If tmp_C2 > 0.0001 Or tmp_C3 <= 0 Or tmp_C2 > 0.1 Then       ' 収束エラー時の処理
            With Sheets(DataSheetName)
                .Range("ZY1").Value = phi_i                        ' Phi_i : φ2in
                .Range("ZZ1").Value = phi_i + pi * 0.92             ' Phi_o : φ2out Initial value
            End With
        Call Calc_Solve_thickness

    End If

    If tmp_C2 > 0.0001 Then
        Debug.Print "   　　差 = " & tmp_C1 & " deg"
        Debug.Print "     誤差 = " & tmp_C2; "":
        Stop
    End If

    '--------------------------
    '--  収束結果の処理
    '--------------------------'

        With Sheets(DataSheetName)
            phi_iw(jj_a) = .Range("ZY1").Value                         ' Phi_i : φ2in
            phi_ow(jj_a) = .Range("ZZ1").Value                         ' Phi_o : φ2out Initial value
            gzai_fs(jj_a) = .Range("AAA4").Value               ' Gzai_fs　ξ'は、内接円中心と基本代数螺旋との距離
        End With

        tmp_C1 = (phi_ow(jj_a) - phi_iw(jj_a)) / pi * 180                     ' 差
        tmp_C2 = Sheets(DataSheetName).Range("AAA1").Value      ' 収束誤差
        tmp_C3 = Sheets(DataSheetName).Range("ZZ1").Value

        Debug.Print "[" & jj_a & "] " & "  Phi_i = " & phi_i * 180 / pi & " deg"

    '        Debug.Print "[W-0] jj_a = " & jj_a & ""
    '        Debug.Print "  Phi_i = " & Phi_i * 180 / PI & " deg"
    '        Debug.Print "  Phi_o = " & Phi_o * 180 / PI & " deg"
    '        Debug.Print "   　　差 = " & tmp_C1 & " deg"
    '        Debug.Print "     誤差 = " & tmp_C2; "":


    '--------------------------
    '　　　内外の基本代数螺旋の偏角φ2(Phi_i, Phi_o)から 内壁包絡線と内接円の交点を求める
    '　　　FS側交点　(xfi,yfi) (xfo,yfo)
    '　　　OS側交点　(xmi,ymi) (xmo,ymo)
    '--------------------------

        phi_i = phi_iw(jj_a)                        ' Phi_i : φ2in
        Phi_o = phi_ow(jj_a)

    ' Formura (8) fi
        Call Wp_xyfi(phi_i)
            w_xfi(jj_a) = Wp_xfi
            w_yfi(jj_a) = Wp_yfi

    ' Formura (7) fo
        Call Wp_xyfo(Phi_o)
            w_xfo(jj_a) = Wp_xfo
            w_yfo(jj_a) = Wp_yfo

    ' Formura (6) mi
        Call Wp_xymi(phi_i)
            w_xmi(jj_a) = Wp_xmi
            w_ymi(jj_a) = Wp_ymi

    ' Formura (5) mo
        Call Wp_xymo(Phi_o)
            w_xmo(jj_a) = Wp_xmo
            w_ymo(jj_a) = Wp_ymo


        thick_fs(jj_a) = Sqr((w_xfo(jj_a) - w_xfi(jj_a)) ^ 2 + (w_yfo(jj_a) - w_yfi(jj_a)) ^ 2)
        thick_os(jj_a) = Sqr((w_xmo(jj_a) - w_xmi(jj_a)) ^ 2 + (w_ymo(jj_a) - w_ymi(jj_a)) ^ 2)

    '-------------
        J = 27                   '【Cellに書き出し用　開始列の番号】
    '-------------

            With Sheets(DataSheetName)
             J = 27:     .Cells(jj_a + 3, J).Value = jj_a                      '【Cellに書き出し】
             J = J + 1:  .Cells(jj_a + 3, J).Value = phi_i                            '【Cellに書き出し】
             J = J + 1:  .Cells(jj_a + 3, J).Value = Phi_o                            '【Cellに書き出し】
             J = J + 1:  .Cells(jj_a + 3, J).Value = .Range("AAA2").Value             '"Gzai_fs"【Cellに書き出し】

             J = J + 1:  .Cells(jj_a + 3, J).Value = thick_fs(jj_a)                     '"Thickness"【Cellに書き出し】
             J = J + 1:  .Cells(jj_a + 3, J).Value = thick_os(jj_a)                      '"Thickness"【Cellに書き出し】
             J = J + 1:  .Cells(jj_a + 3, J).Value = w_xfi(jj_a)
             J = J + 1:  .Cells(jj_a + 3, J).Value = w_yfi(jj_a)
             J = J + 1:  .Cells(jj_a + 3, J).Value = w_xfo(jj_a)
             J = J + 1:  .Cells(jj_a + 3, J).Value = w_yfo(jj_a)
             J = J + 1:  .Cells(jj_a + 3, J).Value = w_xmi(jj_a)
             J = J + 1:  .Cells(jj_a + 3, J).Value = w_ymi(jj_a)
             J = J + 1:  .Cells(jj_a + 3, J).Value = w_xmo(jj_a)
             J = J + 1:  .Cells(jj_a + 3, J).Value = w_ymo(jj_a)

            End With

  Next jj_a

            With Sheets(DataSheetName)
             J = 27:     .Cells(2, J).Value = "Index No."              '
             J = J + 1:  .Cells(2, J).Value = "Phi_i"                  '
             J = J + 1:  .Cells(2, J).Value = "Phi_o"
             J = J + 1:  .Cells(2, J).Value = "Gzai_fs"
             J = J + 1:  .Cells(2, J).Value = "Thickness_FS"
             J = J + 1:  .Cells(2, J).Value = "Thickness_OS"

             J = J + 1:  .Cells(2, J).Value = "w_xfi(i)"
             J = J + 1:  .Cells(2, J).Value = "w_yfi(i)"
             J = J + 1:  .Cells(2, J).Value = "w_xfo(i)"
             J = J + 1:  .Cells(2, J).Value = "w_yfo(i)"
             J = J + 1:  .Cells(2, J).Value = "w_xmi(i)"
             J = J + 1:  .Cells(2, J).Value = "w_ymi(i)"
             J = J + 1:  .Cells(2, J).Value = "w_xmo(i)"
             J = J + 1:  .Cells(2, J).Value = "w_ymo(i)"

            End With

ElseIf Flag_n1 = 0 Then

        Debug.Print "Wrap calc / Flag_n1 = 0 / Data from Cell "

    '-------------
    '    【Cellから読込み】
    '-------------

        J = 5                   '【Cell開始列の番号】

        For jj_a = 0 To tmp_n1                  'Index φ２偏角範囲、圧縮開始＝0～終了tmp_n1

            With Sheets(DataSheetName)
             J = 28:     phi_iw(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  phi_ow(jj_a) = .Cells(jj_a + 3, J).Value      '【Cellから読込み】  ' Phi_i : φ2in
             J = J + 1:  gzai_fs(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  thick_fs(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  thick_fs(jj_a) = .Cells(jj_a + 3, J).Value

             J = J + 1:  w_xfi(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  w_yfi(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  w_xfo(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  w_yfo(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  w_xmi(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  w_ymi(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  w_xmo(jj_a) = .Cells(jj_a + 3, J).Value
             J = J + 1:  w_ymo(jj_a) = .Cells(jj_a + 3, J).Value

            End With

            Debug.Print "[W-F" & jj_a & "] ",
'            Debug.Print "[W-F" & jj_a & "] " & "  Phi_i = " & phi_iw(jj_a) * 180 / PI & " deg"

        Next jj_a

End If

End Sub




'======================================================== 【M5-1-1】
'　関数：　Solver実行、
'　　　　内側代数螺旋の偏角φiから内接円の外側代数螺旋の偏角φoと、内外の代数螺旋の内接円半径ξbを求める
'　　　　内外包絡線の内接円の半径ξfs、ξosを求める
'
'========================================================

Sub Calc_Solve_thickness()
''
''    Sheets("DataSheet_4").Select
'
    SolverOk SetCell:="$AAA$1", MaxMinVal:=2, ValueOf:=0, ByChange:="$ZZ$1", Engine _
        :=1, EngineDesc:="GRG Nonlinear"

    SolverSolve UserFinish:=True

End Sub


'Fn1=================================================== 【M5-1-2】 For Wrap thickness
Sub Wp_xyfi(phi_i As Double)
    ' Formura (8) fi
        Wp_xfi = a * phi_i ^ k * Cos(phi_i - qq) + g1 * Cos(phi_i - qq - Atn(k / phi_i)) + dx
        Wp_yfi = a * phi_i ^ k * Sin(phi_i - qq) + g1 * Sin(phi_i - qq - Atn(k / phi_i)) + dy
End Sub

'Fn2=================================================== 【M5-1-3】
Sub Wp_xyfo(Phi_o As Double)
   ' Formura (7) fo
        Wp_xfo = -a * Phi_o ^ k * Cos(Phi_o - qq) + g1 * Cos(Phi_o - qq - Atn(k / Phi_o)) + dx
        Wp_yfo = -a * Phi_o ^ k * Sin(Phi_o - qq) + g1 * Sin(Phi_o - qq - Atn(k / Phi_o)) + dy
End Sub

'Fn3=================================================== 【M5-1-4】
Sub Wp_xymo(phi_i As Double)
    ' Formura (5) mo   '(-21)
        Wp_xmo = a * phi_i ^ k * Cos(phi_i - qq) - g2 * Cos(phi_i - qq - Atn(k / phi_i)) + dx
        Wp_ymo = a * phi_i ^ k * Sin(phi_i - qq) - g2 * Sin(phi_i - qq - Atn(k / phi_i)) + dy
End Sub

'Fn4=================================================== 【M5-1-5】
Sub Wp_xymi(Phi_o As Double)
    ' Formura (6) mi   '(-14)
        Wp_xmi = -a * Phi_o ^ k * Cos(Phi_o - qq) - g2 * Cos(Phi_o - qq - Atn(k / Phi_o)) + dx
        Wp_ymi = -a * Phi_o ^ k * Sin(Phi_o - qq) - g2 * Sin(Phi_o - qq - Atn(k / Phi_o)) + dy

End Sub





'======================================================== 【M6-1】
'  < Call Calc_GasForce_Ft >
'　　OSに作用する Gas Force Ft：接線方向
'　　　接点(x､y)座標
'      the_c(), the(i), phi1(i), phi2(i),
'                       V_a(i), V_b(i),
'                       Press_A0(i), Press_B0(i)
'                       Press_A1(i), Press_B1(i)
'
'========================================================

Public Sub Calc_GasForce_Ft()   '　(ByVal Phi_1 As Double, ByVal Phi_2 As Double)
    Dim I As Long, J As Long
    Dim jc As Long
    Dim w_n As Long

    Dim tmp As Double

   w_n = turn_wrap_n

' --- Tangensial 接線方向 作用範囲 -----------------------------------------------------------
   ReDim LLt_dxy(dw_n_PI(2))
   ReDim Lt_A(w_n, dw_n_PI(2)):    ReDim Lt_B(w_n, dw_n_PI(2)):     ReDim Lt_D(dw_n_PI(2))
   ReDim LLt_A(w_n, dw_n_PI(2)):   ReDim LLt_B(w_n, dw_n_PI(2))

      ' alpha_dxy :' Angle of OS wrap center offset
      ' RR_dxy    :' Length of OS wrap center offset
      ' LLt_dxy() :' Tangensial length of OS wrap center offset
      ' LLr_dxy() :' Radial length of OS wrap center offset

' --- 接線方向Gas力  -----------------------------------------------------------
   ReDim Ft_a(w_n, dw_n_PI(2)):    ReDim Ft_b(w_n, dw_n_PI(2)):     ReDim Ft_d(dw_n_PI(2))
   ReDim Ft_AB(dw_n_PI(2))

' --- Gas力 Moment -----------------------------------------------------------20171215
'　　接線方向Gas力Moment  Mmt
   ReDim Lmt_A(w_n, dw_n_PI(2)):   ReDim Lmt_B(w_n, dw_n_PI(2)):    ReDim Lmt_D(dw_n_PI(2))
   ReDim Lt_AB(dw_n_PI(2))
   ReDim Mmt_A(w_n, dw_n_PI(2)):   ReDim Mmt_B(w_n, dw_n_PI(2)):    ReDim Mmt_D(dw_n_PI(2))
   ReDim Mmt_AB(dw_n_PI(2))

'------------------
'  作用範囲
'------------------

 For J = 0 To dw_n_PI(2)  ' jc : Theta Index (Rotation angle: 0 ～2*PI)

  'Off set of OS Wrap Center
        LLt_dxy(J) = RR_dxy * Cos(the(J) - qq - alpha_dxy)

  'A,B室　接点間距離  j:外周側から圧縮室No      /  Center Off set : L_dxy * Cos(alpha_dxy)
    For I = 0 To w_n
         ' OS Wrap-xy cordinate
         Lt_A(I, J) = RR_mo_c(I, J) * Cos(beta_fi_c(I, J) + del_mo_c(I, J))         ' Beta:fi, Del:mo
         Lt_B(I, J) = RR_mi_c(I, J) * Cos(beta_mi_c(I, J) + del_mi_c(I, J) + pi)    ' Beta:mi, Del:mi
    Next I
  Next J


 For J = 0 To dw_n_PI(2)
    For I = 0 To w_n - 1

      '-------
      ' A室側　接点間*A室圧力
         If Lt_A(I + 1, J) <> 0 Then

            LLt_A(I, J) = (Lt_A(I, J) - Lt_A(I + 1, J))
            Ft_a(I, J) = LLt_A(I, J) * (Press_A(I + 1, J) - P_suction) * Hw      '[N]=[MPa]*[mm]*[mm]

            'Moment (-)
            Lmt_A(I, J) = ((Lt_A(I, J) + Lt_A(I + 1, J)) / 2 + LLt_dxy(J)) * (-1)
            Mmt_A(I, J) = Ft_a(I, J) * Lmt_A(I, J) / 1000                        '[Nm]=[N]*[mm/1000]

         Else
            LLt_A(I, J) = 0
            Ft_a(I, J) = 0

            Lmt_A(I, J) = 0
            Mmt_A(I, J) = 0
         End If

      '-------
      ' B室側　接点間*B室圧力
         If Lt_B(I, J) <> 0 And Lt_B(I + 1, J) <> 0 Then

            LLt_B(I, J) = (Lt_B(I + 1, J) - Lt_B(I, J))
            Ft_b(I, J) = LLt_B(I, J) * (Press_B(I + 1, J) - P_suction) * Hw

            'Moment (+)
            Lmt_B(I, J) = ((Lt_B(I, J) + Lt_B(I + 1, J)) / 2 + LLt_dxy(J)) * (-1)
            Mmt_B(I, J) = Ft_b(I, J) * Lmt_B(I, J) / 1000                           '[Nm]

         Else
            LLt_B(I, J) = 0
            Ft_b(I, J) = 0

            Lmt_B(I, J) = 0
            Mmt_B(I, J) = 0
         End If

    Next I
 Next J


 For J = 0 To dw_n_PI(2)
    For I = 0 To w_n - 1

      '-------
      ' 吐出室　A,B接点間*吐出圧力
        If Lt_A(I + 1, J) = 0 And Lt_B(I + 1, J) = 0 Then

            Lt_D(J) = (Lt_A(I, J) - Lt_B(I, J))                               '[mm]
            Ft_d(J) = Lt_D(J) * (P_discharge - P_suction) * Hw                '[N]=[mm^2*MPa]

            Lmt_D(J) = ((Lt_A(I, J) + Lt_B(I, J)) / 2 + LLt_dxy(J)) * (-1)    '[mm]
            Mmt_D(J) = Lmt_D(J) * Ft_d(J) / 1000                              '[Nm]

         ElseIf Lt_A(I + 1, J) = 0 Then
            Lt_D(J) = (Lt_A(I, J) - Lt_B(I + 1, J))
            Ft_d(J) = Lt_D(J) * (P_discharge - P_suction) * Hw

            Lmt_D(J) = ((Lt_A(I, J) + Lt_B(I, J)) / 2 + LLt_dxy(J)) * (-1)
            Mmt_D(J) = Lmt_D(J) * Ft_d(J) / 1000

         ElseIf Lt_B(I + 1, J) = 0 Then
            Lt_D(J) = (Lt_A(I + 1, J) - Lt_B(I, J))
            Ft_d(J) = Lt_D(J) * (P_discharge - P_suction) * Hw

            Lmt_D(J) = ((Lt_A(I + 1, J) + Lt_B(I, J)) / 2 + LLt_dxy(J)) * (-1)
            Mmt_D(J) = Lmt_D(J) * Ft_d(J) / 1000

         Else
            Lt_D(J) = (Lt_A(I + 1, J) - Lt_B(I + 1, J))
            Ft_d(J) = Lt_D(J) * (P_discharge - P_suction) * Hw

            Lmt_D(J) = ((Lt_A(I + 1, J) + Lt_B(I + 1, J)) / 2 + LLt_dxy(J)) * (-1)
            Mmt_D(J) = Lmt_D(J) * Ft_d(J) / 1000

        End If

     Next I
  Next J


  For J = 0 To dw_n_PI(2)
       Ft_AB(J) = 0
       Mmt_AB(J) = 0

    For I = 0 To w_n

      '-------
      ' 接線方向Gas Force 合計 Ft
         Ft_AB(J) = Ft_AB(J) + Ft_a(I, J) + Ft_b(I, J)
         Mmt_AB(J) = Mmt_AB(J) + Mmt_A(I, J) + Mmt_B(I, J)

    Next I
         Ft_AB(J) = Ft_AB(J) + Ft_d(J)             '[N]
         Mmt_AB(J) = Mmt_AB(J) + Mmt_D(J)          '[Nm]
         Lt_AB(J) = Mmt_AB(J) / Ft_AB(J) * 1000    '[mm]
  Next J


'Stop

End Sub


'======================================================== 【M6-2】
' < Calc_GasForce_Fr() >
'　　OSに作用する Gas Force Fr：半径方向
'　　　接点(x､y)座標
'      the_c(), the(i), phi1(i), phi2(i),
'                       V_a(i), V_b(i),
'                       Press_A(i,j), Press_B(i,j)
'
'
'========================================================

Public Sub Calc_GasForce_Fr()   '　(ByVal Phi_1 As Double, ByVal Phi_2 As Double)
    Dim I As Long, J As Long
    Dim jc As Long
    Dim w_n As Long

    Dim tmp As Double, tmp_n1 As Double

   w_n = turn_wrap_n

' --- Radial  半径方向 作用範囲-----------------------------------------------------------
   ReDim LLr_dxy(dw_n_PI(2))
   ReDim LLr_A(w_n, dw_n_PI(2)):       ReDim LLr_B(w_n, dw_n_PI(2))
   ReDim Lr_A(w_n, dw_n_PI(2)):        ReDim Lr_B(w_n, dw_n_PI(2)):   ReDim Lr_D(dw_n_PI(2))

      ' alpha_dxy :' Angle of OS wrap center offset
      ' RR_dxy    :' Length of OS wrap center offset
      ' LLt_dxy() :' Tangensial length of OS wrap center offset
      ' LLr_dxy() :' Radial length of OS wrap center offset

' --- 　半径方向 Gas力--------------------------------------------------------------
   ReDim Fr_A(w_n, dw_n_PI(2)):        ReDim Fr_B(w_n, dw_n_PI(2)):   ReDim Fr_D(dw_n_PI(2))
   ReDim Fr_AB(dw_n_PI(2))


' --- Gas力 Moment------------------------------------------------------------------
'　　半径方向 Gas力Moment  Mmt
   ReDim Lmr_A(w_n, dw_n_PI(2)):   ReDim Lmr_B(w_n, dw_n_PI(2)):    ReDim Lmr_D(dw_n_PI(2))
   ReDim Lr_AB(dw_n_PI(2))
   ReDim Mmr_A(w_n, dw_n_PI(2)):   ReDim Mmr_B(w_n, dw_n_PI(2)):    ReDim Mmr_D(dw_n_PI(2))
   ReDim Mmr_AB(dw_n_PI(2))


'------------------
'  作用範囲
'------------------
'

 For J = 0 To dw_n_PI(2)  ' jc : Theta Index (Rotation angle: 0 ～2*PI)

         LLr_dxy(J) = RR_dxy * Sin(the(J) - qq - alpha_dxy)
                        ' RR_dxy    : Length of OS wrap center offset
                        ' alpha_dxy : Angle of OS wrap center offset
    For I = 0 To w_n
         Lr_A(I, J) = RR_mo_c(I, J) * Sin(beta_fi_c(I, J) + del_mo_c(I, J))         ' LLr_dxy(j)
         Lr_B(I, J) = RR_mi_c(I, J) * Sin(beta_mi_c(I, J) + del_mi_c(I, J) + pi)    ' LLr_dxy(j)
    Next I

  Next J

  For J = 0 To dw_n_PI(2)
    For I = 0 To w_n - 1

      '-------
      'A室側　接点間*A室圧力
         If Lr_A(I + 1, J) <> 0 Then

            LLr_A(I, J) = (Lr_A(I + 1, J) - Lr_A(I, J))                       'Abs   A: 2A-1A
'            Fr_A(i, j) = LLr_A(i, j) * (P_discharge - Press_A(i + 1, j) ) * Hw
            Fr_A(I, J) = LLr_A(I, J) * (P_discharge - Press_A(I + 1, J) + Press_A(I, J) - P_suction) * Hw

            'Moment (-)
            Lmr_A(I, J) = ((Lr_A(I, J) + Lr_A(I + 1, J)) / 2 + LLr_dxy(J)) * (1)    '
            Mmr_A(I, J) = Fr_A(I, J) * Lmr_A(I, J) / 1000
         Else
            LLr_A(I, J) = 0
            Fr_A(I, J) = 0

            Lmr_A(I, J) = 0
            Mmr_A(I, J) = 0
         End If

      '-------
      'B室側　接点間*B室圧力
         If Lr_B(I, J) <> 0 And Lr_B(I + 1, J) <> 0 Then

            LLr_B(I, J) = (Lr_B(I, J) - Lr_B(I + 1, J))                             'Abs   B: -(2B-1B)
            Fr_B(I, J) = LLr_B(I, J) * (P_discharge - Press_B(I + 1, J)) * Hw

            'Moment (+)
            Lmr_B(I, J) = ((Lr_B(I, J) + Lr_B(I + 1, J)) / 2 + LLr_dxy(J)) * (1)    'Abs
            Mmr_B(I, J) = Fr_B(I, J) * Lmr_B(I, J) / 1000
          Else
            LLr_B(I, J) = 0
            Fr_B(I, J) = 0

            Lmr_B(I, J) = 0
            Mmr_B(I, J) = 0
         End If

    Next I
 Next J


  For J = 0 To dw_n_PI(2)
   '    tmp_n1  = w_n
       tmp_n1 = 2

      '-------
      '吐出室　接点間*吐出圧力
'            Lr_D(j) = (Lr_A(0, j) - Lr_B(0, j))
            Lr_D(J) = (Lr_A(tmp_n1, J) - Lr_B(tmp_n1, J))
            Fr_D(J) = Lr_D(J) * (P_discharge - P_suction) * Hw

            Lmr_D(J) = ((Lr_A(tmp_n1, J) + Lr_B(tmp_n1, J)) / 2 + LLr_dxy(J)) * (1)
            Mmr_D(J) = Lmr_D(J) * Fr_D(J) / 1000                     ' [Nm]

    '-------
    '吐出室　接点間*吐出圧力
    '         If Lr_A(i + 1, j) = 0 And Lr_B(i + 1, j) = 0 Then
    '
    '            Lr_D(j) = (Lr_A(i, j) - Lr_B(i, j))
    '            Fr_D(j) = Lr_D(j) * (P_discharge - P_suction) * Hw
    '
    '            Lmr_D(j) = ((Lr_A(i, j) + Lr_B(i, j)) / 2 + LLr_dxy(j)) * (-1)
    '            Mmr_D(j) = Lmr_D(j) * Fr_D(j) / 1000
    '
    '         ElseIf Lr_A(i + 1, j) = 0 Then
    '            Lr_D(j) = (Lr_A(i, j) - Lr_B(i + 1, j))
    '            Fr_D(j) = Lr_D(j) * (P_discharge - P_suction) * Hw
    '
    '            Lmr_D(j) = ((Lr_A(i, j) + Lr_B(i, j)) / 2 + LLr_dxy(j)) * (-1)
    '            Mmr_D(j) = Lmr_D(j) * Fr_D(j) / 1000
    '
    '         ElseIf Lr_B(i + 1, j) = 0 Then
    '            Lr_D(j) = (Lr_A(i + 1, j) - Lr_B(i, j))
    '            Fr_D(j) = Lr_D(j) * (P_discharge - P_suction) * Hw
    '
    '            Lmr_D(j) = ((Lr_A(i + 1, j) + Lr_B(i, j)) / 2 + LLr_dxy(j)) * (-1)
    '            Mmr_D(j) = Lmr_D(j) * Fr_D(j) / 1000
    '
    '         Else
    '            Lr_D(j) = (Lr_A(i + 1, j) - Lr_B(i + 1, j))
    '            Fr_D(j) = Lr_D(j) * (P_discharge - P_suction) * Hw
    '
    '            Lmr_D(j) = ((Lr_A(i + 1, j) + Lr_B(i + 1, j)) / 2 + LLr_dxy(j)) * (-1)
    '            Mmr_D(j) = Lmr_D(j) * Fr_D(j) / 1000
    '
    '         End If


 Next J


  For J = 0 To dw_n_PI(2)
      Fr_AB(J) = 0
      Mmr_AB(J) = 0

    For I = 0 To w_n

      '-------
      '法線方向Gas Force 合計 Fr
         Fr_AB(J) = Fr_AB(J) + Fr_A(I, J) + Fr_B(I, J)         ' [N]
         Mmr_AB(J) = Mmr_AB(J) + Mmr_A(I, J) + Mmr_B(I, J)     ' [Nm]

    Next I
        Fr_AB(J) = Fr_AB(J) + Fr_D(J)                          ' [N]
         Mmr_AB(J) = Mmr_AB(J) + Mmr_D(J)                      ' [Nm]
         Lr_AB(J) = Mmr_AB(J) / Fr_AB(J) * 1000                ' [mm]   γ= R_Fgc_r

  Next J


'Stop

End Sub


'======================================================== 【M6-1】
' < Calc_GasForce_Fz() >
'　　OSに作用する Gas Force Ft：接線方向
'　　　接点(x､y)座標
'      the_c(), the(i), phi1(i), phi2(i),
'                       V_a(i), V_b(i),
'                       Press_A0(i), Press_B0(i)
'                       Press_A1(i), Press_B1(i)
'
'========================================================

Public Sub Calc_GasForce_Fz()   '　(ByVal Phi_1 As Double, ByVal Phi_2 As Double)
    Dim I   As Long:   Dim J As Long
    Dim jn  As Long:
'    Dim w_n As Long

    Dim tmp     As Double:
    Dim tmp_x  As Double:     Dim tmp_y  As Double:      Dim tmp_r  As Double
    Dim tmp_q1  As Double:    Dim tmp_q2  As Double:     Dim tmp_q3  As Double

    Dim Pin_tmp As Double:    Dim Pout_tmp As Double  'for area pressuer

' aiba's variants for calculating
   Dim Pp(4) As Double:
   Dim Pg(4) As Double:
   '    Dim Pp(1) As Double:    Dim Pp(2) As Double:   Dim Pp(3) As Double  ' A chamber pressure
   '    Dim Pg(1) As Double:    Dim Pg(2) As Double:   Dim Pg(3) As Double  ' B chamber pressure
   Dim Pw As Double:    'Dim P_mid As Double:


    jn = 11
    J = jn

' --- 作用Gas圧力 ---------------------------------------------
   ReDim Pz_A(J, dw_n) As Double:  ReDim Pz_B(J, dw_n) As Double
   ReDim Pz_f(J, dw_n) As Double:  ReDim Pz_m(J, dw_n) As Double
   ReDim Pz_Za(J, dw_n) As Double:  ReDim Pz_Zb(J, dw_n) As Double

   ReDim DP_tm(J, dw_n) As Double:  ReDim DP_tf(J, dw_n) As Double    'Difference pressure of Wrap tips

' --- 軸方向Gas力  -----------------------------------------------------------
   ReDim Fz_A(J, dw_n) As Double:  ReDim Fz_B(J, dw_n) As Double:
   ReDim Fz_f(J, dw_n) As Double:  ReDim Fz_m(J, dw_n) As Double:
   ReDim Fz_Za(dw_n) As Double: ReDim Fz_Zb(dw_n) As Double:
   ReDim Fz_sp(dw_n) As Double:

' --- Gas力 Moment -------------------------------------------------------------
   ReDim Mzx_f(J, dw_n) As Double:   ReDim Mzy_f(J, dw_n) As Double:
   ReDim Mzx_m(J, dw_n) As Double:   ReDim Mzy_m(J, dw_n) As Double:
   ReDim Mzx_A(J, dw_n) As Double:   ReDim Mzy_A(J, dw_n) As Double:
   ReDim Mzx_B(J, dw_n) As Double:   ReDim Mzy_B(J, dw_n) As Double:

   ReDim Mzx_Za(dw_n) As Double:   ReDim Mzy_Za(dw_n) As Double:
   ReDim Mzx_Zb(dw_n) As Double:   ReDim Mzy_Zb(dw_n) As Double:

' --- Gas力 active point ----------
   ReDim x_zforce_a(dw_n) As Double:      ReDim y_zforce_a(dw_n) As Double     ' FS x-y cordinate 20180323
   ReDim x_zforce_b(dw_n) As Double:      ReDim y_zforce_b(dw_n) As Double

   ReDim r_zforce_a(dw_n) As Double:      ReDim t_zforce_a(dw_n) As Double     ' OS r-t cordinate 20180658
   ReDim r_zforce_b(dw_n) As Double:      ReDim t_zforce_b(dw_n) As Double

'---------------------------------------------------------

'GoTo label_Calc_GasForce_Fz_end

'    atm_00 = 0.1013


For I = 0 To 0   ' dw_n   ' 軸回転範囲 4PI = index 366  , 2PI = index 183 , PI = index 90

   ' the_1 = the(i) - qq

      Pp(1) = P_discharge      '= Press_A(3, i)     ' [MPa(abs)]
      Pp(2) = Press_A(2, I)    '= Press_A(2, i)     ' [MPa(abs)]
      Pp(3) = Press_A(1, I)    '= Press_A(1, i)     ' [MPa(abs)]
      Pp(4) = P_suction        '= Press_A(0, i) = P_suction    ' [MPa(abs)]

      Pg(1) = P_discharge      '= Press_B(3, i)     ' [MPa(abs)]
      Pg(2) = Press_B(2, I)    '= Press_B(2, i)     ' [MPa(abs)]
      Pg(3) = P_suction        '= Press_B(1, i)     ' [MPa(abs)]

      Pw = P_groove
      P_mid = P_back                'P_suction + (P_discharge - P_suction) * 0.5

 '--< Wrap Side / Axial Force >------

   '--- A Chamber
   '[Pp(3) : A Chamber Fs_in-OS_out]  Sg2_a(3, i)
      tmp = 3
         Pz_A(3, I) = Pp(3) - atm_00
         Fz_A(3, I) = Sg2_a(3, I) * Pz_A(3, I)

         Mzx_A(tmp, I) = Fz_A(tmp, I) * xg2_a(tmp, I)
         Mzy_A(tmp, I) = Fz_A(tmp, I) * yg2_a(tmp, I)

   '[Pp(2) : A Chamber Fs_in-OS_out]  Sg2_a(2, i)
      tmp = 2
         Pz_A(2, I) = Pp(2) - atm_00
         Fz_A(2, I) = Sg2_a(2, I) * Pz_A(2, I)

         Mzx_A(tmp, I) = Fz_A(tmp, I) * xg2_a(tmp, I)
         Mzy_A(tmp, I) = Fz_A(tmp, I) * yg2_a(tmp, I)

   '--- D Chamber : Discharge
   '[ppi-Pg(1)] Discharge chamber     Sg2_a(0, i) <==={Sg_f(0, i) & Sg_m(0, i)}
   '      Pp(1) = P_discharge
   '      Pg(1) = P_discharge
         Pz_A(0, I) = Pg(1) - atm_00
         Fz_A(0, I) = Sg2_a(0, I) * Pz_A(0, I)
      tmp = 0
         Mzx_A(tmp, I) = Fz_A(tmp, I) * xg2_a(tmp, I)
         Mzy_A(tmp, I) = Fz_A(tmp, I) * yg2_a(tmp, I)

   '--- B Chamber
   '[Pg(3) : Suction Chamber ]        Sg2_b(3, i)
         Pz_B(3, I) = Pg(3) - atm_00
         Fz_B(3, I) = Sg2_b(3, I) * Pz_B(3, I)
      tmp = 3
         Mzx_B(tmp, I) = Fz_B(tmp, I) * xg2_b(tmp, I)
         Mzy_B(tmp, I) = Fz_B(tmp, I) * yg2_b(tmp, I)

   '[Pg(2) : B Chamber OS_in-FS_ou]   Sg2_b(2, i)
         Pz_B(2, I) = Pg(2) - atm_00
         Fz_B(2, I) = Sg2_b(2, I) * Pz_B(2, I)
      tmp = 2
         Mzx_B(tmp, I) = Fz_B(tmp, I) * xg2_b(tmp, I)
         Mzy_B(tmp, I) = Fz_B(tmp, I) * yg2_b(tmp, I)

   '----------
   '[PW : oil groove ]              Sg2_b(7, i)
   '      Pw = P_discharge            'oil groove pressure
         Pz_B(7, I) = Pw - atm_00
         Fz_B(7, I) = Sg2_b(7, I) * Pz_B(7, I)
      tmp = 7
         Mzx_B(tmp, I) = Fz_B(tmp, I) * xg2_b(tmp, I)
         Mzy_B(tmp, I) = Fz_B(tmp, I) * yg2_b(tmp, I)

   '----------
      '[Ml1 : OS inner seal] Sg2_b(8, i)     ==> Seal Ring side
      '[Ml2 : OS_Plate & seal] Sg2_b(9, i)   ==> Seal Ring side


   '----------
   '[tp6 OS]   Sg_m(6, i)
         Pout_tmp = P_suction    '= Press_A(0, i)
         Pin_tmp = P_suction     '= Press_B(0, i)
            Pz_m(6, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_m(6, I) = Sg_m(6, I) * Pz_m(6, I)
         tmp = 6
            DP_tm(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_m(tmp, I) = Fz_m(tmp, I) * xg_m(tmp, I)
            Mzy_m(tmp, I) = Fz_m(tmp, I) * yg_m(tmp, I)

   '[tp5 OS]   Sg_m(5, i)
         Pout_tmp = Pp(3)          'Press_A(1, i)
         Pin_tmp = Pg(3)           'P_suction = Press_B(1, i)
            Pz_m(5, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_m(5, I) = Sg_m(5, I) * Pz_m(5, I)
         tmp = 5
            DP_tm(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_m(tmp, I) = Fz_m(tmp, I) * xg_m(tmp, I)
            Mzy_m(tmp, I) = Fz_m(tmp, I) * yg_m(tmp, I)

   '[tp4 OS]   Sg_m(4, i)
         Pout_tmp = Pp(3)          'Press_A(1, i)   '= Pp(3)
         Pin_tmp = Pg(2)           'Press_B(2, i)   '= Pg(2)
            Pz_m(4, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_m(4, I) = Sg_m(4, I) * Pz_m(4, I)
         tmp = 4
            DP_tm(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_m(tmp, I) = Fz_m(tmp, I) * xg_m(tmp, I)
            Mzy_m(tmp, I) = Fz_m(tmp, I) * yg_m(tmp, I)

   '[tp3 OS]   Sg_m(3, i)
         Pout_tmp = Pp(2)          'Press_A(2, i)   '= Pp(2)
         Pin_tmp = Pg(2)           'Press_B(2, i)   '= Pg(2)
            Pz_m(3, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_m(3, I) = Sg_m(3, I) * Pz_m(3, I)
         tmp = 3
            DP_tm(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_m(tmp, I) = Fz_m(tmp, I) * xg_m(tmp, I)
            Mzy_m(tmp, I) = Fz_m(tmp, I) * yg_m(tmp, I)

   '[tp2 OS]   Sg_m(2, i)
         Pout_tmp = Pp(2)          'Press_A(1, i)   '= Pp(2)
         Pin_tmp = Pg(1)           'discharge '=Press_B(3, i)    '= Pg(1)
            Pz_m(2, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_m(2, I) = Sg_m(2, I) * Pz_m(2, I)
         tmp = 2
            DP_tm(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_m(tmp, I) = Fz_m(tmp, I) * xg_m(tmp, I)
            Mzy_m(tmp, I) = Fz_m(tmp, I) * yg_m(tmp, I)

   '[tp1 OS]   Sg_m(1, i)
         Pout_tmp = Pp(1)          '= discharge
         Pin_tmp = Pg(1)           '= discharge
            Pz_m(1, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_m(1, I) = Sg_m(1, I) * Pz_m(1, I)
         tmp = 1
            DP_tm(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_m(tmp, I) = Fz_m(tmp, I) * xg_m(tmp, I)
            Mzy_m(tmp, I) = Fz_m(tmp, I) * yg_m(tmp, I)

   '-----------
   '[tg1 FS]   Sg_f(1, i)
         Pout_tmp = Pg(1)          '= discharge
         Pin_tmp = Pp(1)           '= discharge
            Pz_f(1, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(1, I) = Sg_f(1, I) * Pz_f(1, I)
         tmp = 1
            DP_tf(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '[tg2 FS]   Sg_f(2, i)
         Pout_tmp = Pg(2)          '= Press_B(0, i)
         Pin_tmp = Pp(1)           '= Press_A(0, i)
            Pz_f(2, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(2, I) = Sg_f(2, I) * Pz_f(2, I)
         tmp = 2
            DP_tf(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '[tg3 FS]   Sg_f(3, i)
         Pout_tmp = Pg(2)          'Press_A(0, i)
         Pin_tmp = Pp(2)           'Press_A(0, i)
            Pz_f(3, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(3, I) = Sg_f(3, I) * Pz_f(3, I)
         tmp = 3
            DP_tf(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '[tg4 FS]   Sg_f(4, i)
         Pout_tmp = Pg(3)          'Press_A(1, i)   'Pg(3) =P_suction
         Pin_tmp = Pp(2)           'Press_B(2, i)    'Pg(2) =P_suction
            Pz_f(4, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(4, I) = Sg_f(4, I) * Pz_f(4, I)
         tmp = 4
            DP_tf(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '-----------
   '[tg5 : FS-Tip /inlet & FSin]    Sg_f(5, i)
   '
         Pout_tmp = Pg(3)          'Press_A(1, i)   'Pg(3) =P_suction
         Pin_tmp = Pp(3)           'Press_B(2, i)    'Pg(2) =P_suction
            Pz_f(5, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(5, I) = Sg_f(5, I) * Pz_f(5, I)
         tmp = 5
            DP_tf(tmp, I) = Pin_tmp - Pout_tmp
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '[tg6 :     Sg_f(6, i)
   ' OS Plate / Suction inlet
         Pout_tmp = P_mid          'Press_A(1, i)   'Pg(3) =P_suction
         Pin_tmp = Pp(3)           'Press_B(2, i)    'Pg(2) =P_suction
            Pz_f(6, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(6, I) = Sg_f(6, I) * Pz_f(6, I)
         tmp = 6
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '[tg7 :     Sg_f(7, i)
   ' oil groove inner arc1 / FS_in
         Pout_tmp = Pw          '
         Pin_tmp = Pp(3)           '
            Pz_f(7, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(7, I) = Sg_f(7, I) * Pz_f(7, I)
         tmp = 7
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '[tg8 :     Sg_f(8, i)
   ' oil-groove inner arc1 / FS_in Suction-Inlet arc (FS_in)
         Pout_tmp = Pw          '
         Pin_tmp = Pg(3)           '
            Pz_f(8, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(8, I) = Sg_f(8, I) * Pz_f(8, I)
         tmp = 8
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '[tg9 :     Sg_f(9, i)
   ' oil groove / Suction inlet
         Pout_tmp = P_mid         '
         Pin_tmp = Pg(3)           '
            Pz_f(9, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(9, I) = Sg_f(9, I) * Pz_f(9, I)
         tmp = 9
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '[tg10 : oil & OS_Plate ]        Sg_f(10, i)
   ' OS_Plate / oil groove
         Pout_tmp = P_mid         '
         Pin_tmp = Pw          '
            Pz_f(10, I) = (Pout_tmp + Pin_tmp) / 2 - atm_00
            Fz_f(10, I) = Sg_f(10, I) * Pz_f(10, I)
         tmp = 10
            Mzx_f(tmp, I) = Fz_f(tmp, I) * xg_f(tmp, I)
            Mzy_f(tmp, I) = Fz_f(tmp, I) * yg_f(tmp, I)

   '-----------
   '[ Fz axial Gas Force : Wrap side
   '-----------
       For J = 0 To jn

         '--Fz axial Gas Force : Wrap side
         Fz_Za(I) = Fz_Za(I) + (Fz_A(J, I) + Fz_B(J, I) + Fz_f(J, I) + Fz_m(J, I))

         '--Fz axial Gas Moment : Wrap side
         Mzx_Za(I) = Mzx_Za(I) + (Mzx_A(J, I) + Mzx_B(J, I) + Mzx_f(J, I) + Mzx_m(J, I))
         Mzy_Za(I) = Mzy_Za(I) + (Mzy_A(J, I) + Mzy_B(J, I) + Mzy_f(J, I) + Mzy_m(J, I))

       Next J

      If Fz_Za(I) = 0 Then
         Stop
      End If

'      Mzx_Za(i) = -50110.11874
'      Mzy_Za(i) = -51405.92202

      '-- Fz axial Force active point on FS-xy cordinate
         x_zforce_a(I) = Mzx_Za(I) / Fz_Za(I)
         y_zforce_a(I) = Mzy_Za(I) / Fz_Za(I)

'         tmp_x = x_zforce_a(i) - x_e(i)          '[mm]  OS-Plate xy cordinate
'         tmp_y = y_zforce_a(i) - y_e(i)
         tmp_r = Sqr((x_zforce_a(I)) ^ 2 + (y_zforce_a(I)) ^ 2)

         tmp_q1 = Atn(y_zforce_a(I) / x_zforce_a(I))
         tmp_q2 = Atn(y_e(I) / x_e(I))

'         tmp_q1 = Atan2(y_zforce_a(i), x_zforce_a(i))
'         tmp_q2 = Atan2(y_e(i), x_e(i))
         tmp_q3 = the(I) - qq

         r_zforce_a(I) = tmp_r * Cos(tmp_q1 - tmp_q2) - Ro   '[mm]  OS-Plate rt cordinate
         t_zforce_a(I) = tmp_r * Sin(tmp_q1 - tmp_q2)


 '--< Seal Ring Side / Axial Force >------

   '-----------
   '[ML-1 : OS_Plate & seal]         Sg2_b(8, i)
         Pz_B(8, I) = P_mid - atm_00
         Fz_B(8, I) = Sg2_b(8, I) * Pz_B(8, I)
      tmp = 8
         Mzx_B(8, I) = Fz_B(tmp, I) * xg2_b(8, I)
         Mzy_B(8, I) = Fz_B(tmp, I) * yg2_b(8, I)

   '-----------
   '[ML-2 : OS inner seal]           Sg2_b(9, i)
         Pz_B(9, I) = (P_discharge) - atm_00
         Fz_B(9, I) = Sg2_b(9, I) * Pz_B(9, I)
      tmp = 9
         Mzx_B(9, I) = Fz_B(tmp, I) * xg2_b(9, I)    '[N-mm]
         Mzy_B(9, I) = Fz_B(tmp, I) * yg2_b(9, I)


   '-----------
   '[OS gravity force]

         Fz_B(10, I) = F_mg                           '[N]
         xg2_b(10, I) = R_Fmg_r                       '[mm]
         yg2_b(10, I) = R_Fmg_t

         Mzx_B(10, I) = Fz_B(10, I) * xg2_b(10, I)    '[N-mm]
         Mzy_B(10, I) = Fz_B(10, I) * yg2_b(10, I)    '[N-mm]


   '-----------
   '[ Fz axial Gas Force/ Moment / point : Seal Ring side
   '-----------
       For J = 8 To 9

         '--Fz axial Gas Force :
         Fz_Zb(I) = Fz_Zb(I) + Fz_B(J, I)       '[N]

         '--Fz axial Gas Moment :
         Mzx_Zb(I) = Mzx_Zb(I) + Mzx_B(J, I)    '[N-mm]
         Mzy_Zb(I) = Mzy_Zb(I) + Mzy_B(J, I)

       Next J

      If Fz_Zb(I) = 0 Then
         Stop
      End If

        '-- Fz axial Force active point xy
         x_zforce_b(I) = Mzx_Zb(I) / Fz_Zb(I)    '[mm]  FS-xy cordinate
         y_zforce_b(I) = Mzy_Zb(I) / Fz_Zb(I)

'         tmp_x = x_zforce_b(i) - x_e(i)          '[mm]  OS-Plate xy cordinate
'         tmp_y = y_zforce_b(i) - y_e(i)
         tmp_r = Sqr((x_zforce_b(I)) ^ 2 + (y_zforce_b(I)) ^ 2)
         tmp_q1 = Atan2(y_zforce_b(I), x_zforce_b(I))
         tmp_q2 = Atan2(y_e(I), x_e(I))
         tmp_q3 = the(I) - qq

         r_zforce_b(I) = tmp_r * Cos(tmp_q1 - tmp_q2) - Ro   '[mm]  OS-Plate rt cordinate
         t_zforce_b(I) = tmp_r * Sin(tmp_q1 - tmp_q2)


   '-----------
   '[ Fsp axial thrust Force/ Moment / point : thrust reflection force
   '-----------

         Fz_sp(I) = Fz_Zb(I) - Fz_Za(I) - F_mg    '[N]

Next I

label_Calc_GasForce_Fz_end:

End Sub



'======================================================== 【    】
'  結果を、配列に保存　　===>    ※ Aiba 結果と照合、検証用　20180627
'                                   注意) 配列共用 Data_Strage_3(dw_c, dw_c)
'　　　　配列を、直接Cellへ格納
'
'　　　　※注）配列とCellの行と列は逆
'========================================================

Public Sub Data_Strage_to_array_2()

Dim I As Long, J As Long
Dim ii As Long, jj As Long

Dim I1 As Long, J1 As Long
Dim dw_n3 As Long

'【Cellに書き出し】

'Index_i

'--------------------------
'-- Data Title 部
'--------------------------


    I = 3     ' 代入する配列の開始行
      dw_n3 = I + 20 + 289 + 8 + 2    ' 　配列の行数
      dw_n3 = dw_n3 + 13              '= i + 332

    J = 1     ' 代入する配列の開始列
      dw_c = 3  ' 　配列の列

    ReDim Data_Strage_3(1 To dw_n3, 1 To dw_c)       ' Data_Strage_3(dw_c, dw_n + 3)
         'Data_Strage_3(i, j)　　　                  ' i=配列の行, j=配列の列

         '参考) Cells(2, 1).Select = Range("A2").Select　　'Cell(3,8) = Range("H3")

 '------ 1-20 [項目名]  A室側 外から１番目の接点

               Data_Strage_3(I, J) = "N_rps"
   I = I + 1:   Data_Strage_3(I, J) = "Hw"
   I = I + 1:   Data_Strage_3(I, J) = "P_discharge_(G)"
   I = I + 1:   Data_Strage_3(I, J) = "P_suction_(G)"
   I = I + 1:   Data_Strage_3(I, J) = "OS_seal"
   I = I + 1:   Data_Strage_3(I, J) = "Ro"
   I = I + 1:   Data_Strage_3(I, J) = "delta_ky"
   I = I + 1:   Data_Strage_3(I, J) = "alpha_ky"
   I = I + 1:   Data_Strage_3(I, J) = "myu_sb"
   I = I + 1:   Data_Strage_3(I, J) = "myu_ky"
   I = I + 1:   Data_Strage_3(I, J) = "myu_th"
   I = I + 1:   Data_Strage_3(I, J) = "h_pl"
   I = I + 1:   Data_Strage_3(I, J) = "OS_dia"
   I = I + 1:   Data_Strage_3(I, J) = "h_ky"
   I = I + 1:   Data_Strage_3(I, J) = "b_kos"
   I = I + 1:   Data_Strage_3(I, J) = "b_kmf"
   I = I + 1:   Data_Strage_3(I, J) = "Z_eb"
   I = I + 1:   Data_Strage_3(I, J) = "R_kos"
   I = I + 1:   Data_Strage_3(I, J) = "R_kmf"
   I = I + 1:   Data_Strage_3(I, J) = "Fs_r"
   I = I + 1:   Data_Strage_3(I, J) = "Index_i"
   I = I + 1:   Data_Strage_3(I, J) = "N_wrap_a"
   I = I + 1:   Data_Strage_3(I, J) = "N_wrap_b"
   I = I + 1:   Data_Strage_3(I, J) = "the_c"
   I = I + 1:   Data_Strage_3(I, J) = "Ros_x=Ro * Cos(the())"
   I = I + 1:   Data_Strage_3(I, J) = "Ros_y = Ro * Sin(the())"
   I = I + 1:   Data_Strage_3(I, J) = "pp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "pp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "pp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "pp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "pg(1)"
   I = I + 1:   Data_Strage_3(I, J) = "pg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "pg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "tp1"
   I = I + 1:   Data_Strage_3(I, J) = "tp2"
   I = I + 1:   Data_Strage_3(I, J) = "tp3"
   I = I + 1:   Data_Strage_3(I, J) = "tp4"
   I = I + 1:   Data_Strage_3(I, J) = "tp5"
   I = I + 1:   Data_Strage_3(I, J) = "tp6"
   I = I + 1:   Data_Strage_3(I, J) = "tg1"
   I = I + 1:   Data_Strage_3(I, J) = "tg2"
   I = I + 1:   Data_Strage_3(I, J) = "tg3"
   I = I + 1:   Data_Strage_3(I, J) = "tg4"
   I = I + 1:   Data_Strage_3(I, J) = "tg5"
   I = I + 1:   Data_Strage_3(I, J) = "tg6"
   I = I + 1:   Data_Strage_3(I, J) = "tg7"
   I = I + 1:   Data_Strage_3(I, J) = "tg8"
   I = I + 1:   Data_Strage_3(I, J) = "tg9"
   I = I + 1:   Data_Strage_3(I, J) = "tg10"
   I = I + 1:   Data_Strage_3(I, J) = "tg11"
   I = I + 1:   Data_Strage_3(I, J) = "P_groove"
   I = I + 1:   Data_Strage_3(I, J) = "P_back"
   I = I + 1:   Data_Strage_3(I, J) = "Fgc_r"
   I = I + 1:   Data_Strage_3(I, J) = "Fgc_t"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fgc_t"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fgc_r"
   I = I + 1:   Data_Strage_3(I, J) = "Mst"
   I = I + 1:   Data_Strage_3(I, J) = "Msr"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tp(6)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tp(6)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tp(6)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(6)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tp(6)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tp(6)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tp(6)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tp(5)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tp(5)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tp(5)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(5)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tp(5)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tp(5)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tp(5)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tp(4)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tp(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tp(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tp(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(6)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(6)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(6)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(6)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(6)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(6)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(6)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(5)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(5)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(5)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(5)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(5)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(5)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(5)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(4)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(4)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(4)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(4)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(4)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(4)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(4)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(1)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(1)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(1)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(1)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(1)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(7)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(7)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(7)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(7)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(7)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(7)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(7)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(8)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(8)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(8)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(8)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(8)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(8)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(8)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(9)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(9)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(9)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(9)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(9)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(9)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(9)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(10)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(10)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(10)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(10)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(10)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(10)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(10)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(11)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(11)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(11)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(11)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(11)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(11)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(11)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_tg(12)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_tg(12)"
   I = I + 1:   Data_Strage_3(I, J) = "S_tg(12)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(12)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_tg(12)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_tg(12)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_tg(12)"
   I = I + 1:   Data_Strage_3(I, J) = "Fgc_z"         '[227]
   I = I + 1:   Data_Strage_3(I, J) = "M_x_Fa"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_Fa"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_Fa"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_Fa"
   I = I + 1:   Data_Strage_3(I, J) = "Ros_x_Fa"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_Fa"
   I = I + 1:   Data_Strage_3(I, J) = "Ros_Fgz_r"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fgz_t"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pb(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pb(1)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pb(1)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pb(1)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pb(1)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pb(1)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pb(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pb(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pb(2)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pb(2)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pb(2)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pb(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pb(2)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pb(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_pb(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_pb(3)"
   I = I + 1:   Data_Strage_3(I, J) = "S_pb(3)"
   I = I + 1:   Data_Strage_3(I, J) = "DP_pb(3)"
   I = I + 1:   Data_Strage_3(I, J) = "Fz_pb(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_x_pb(3)"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_pb(3)"
   I = I + 1:   Data_Strage_3(I, J) = "Fgb_z"        '[257]
   I = I + 1:   Data_Strage_3(I, J) = "M_x_Fgb"
   I = I + 1:   Data_Strage_3(I, J) = "M_y_Fgb"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_Fgb"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_Fgb"
   I = I + 1:   Data_Strage_3(I, J) = "Ros_x_Fgb"
   I = I + 1:   Data_Strage_3(I, J) = "Ros_y_Fgb"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fgb_r"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fgb_t"
   I = I + 1:   Data_Strage_3(I, J) = "Fsp_z"

   I = I + 1:   Data_Strage_3(I, J) = "R_x_osw(1)"    '[267]
   I = I + 1:   Data_Strage_3(I, J) = "R_y_osw(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_z_osw(1)"
   I = I + 1:   Data_Strage_3(I, J) = "V_osw(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_osw(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_osw(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_z_osw(2)"
   I = I + 1:   Data_Strage_3(I, J) = "V_osw(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_osw(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_osw(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_z_osw(3)"
   I = I + 1:   Data_Strage_3(I, J) = "V_osw(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_osw(0)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_osw(0)"
   I = I + 1:   Data_Strage_3(I, J) = "R_z_osw(0)"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fmg_r"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fmg_t"
   I = I + 1:   Data_Strage_3(I, J) = "vol_os"
   I = I + 1:   Data_Strage_3(I, J) = "m_os"
   I = I + 1:   Data_Strage_3(I, J) = "F_mg"
   I = I + 1:   Data_Strage_3(I, J) = "Fmc_r"

   I = I + 1:   Data_Strage_3(I, J) = "Ros_F1_oy"   '[288]
   I = I + 1:   Data_Strage_3(I, J) = "Ros_F2_oy"
   I = I + 1:   Data_Strage_3(I, J) = "Ros_F3_ox"
   I = I + 1:   Data_Strage_3(I, J) = "Ros_F4_ox"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_or(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_or(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_z_or(1)"
   I = I + 1:   Data_Strage_3(I, J) = "V_or(1)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_or(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_or(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_z_or(2)"
   I = I + 1:   Data_Strage_3(I, J) = "V_or(2)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_or(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_or(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_z_or(3)"
   I = I + 1:   Data_Strage_3(I, J) = "V_or(3)"
   I = I + 1:   Data_Strage_3(I, J) = "R_x_or(0)"
   I = I + 1:   Data_Strage_3(I, J) = "R_y_or(0)"
   I = I + 1:   Data_Strage_3(I, J) = "R_z_or(0)"
   I = I + 1:   Data_Strage_3(I, J) = "vol_or"
   I = I + 1:   Data_Strage_3(I, J) = "m_or"
   I = I + 1:   Data_Strage_3(I, J) = "Fc_or"      '[309]

  ' Matrix
   I = I + 1:   Data_Strage_3(I, J) = "Fk_1"
   I = I + 1:   Data_Strage_3(I, J) = "Fk_2"
   I = I + 1:   Data_Strage_3(I, J) = "Fk_3"
   I = I + 1:   Data_Strage_3(I, J) = "Fk_4"
   I = I + 1:   Data_Strage_3(I, J) = "Fsb_t"
   I = I + 1:   Data_Strage_3(I, J) = "Fsb_r"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fsp_t :Ct"
   I = I + 1:   Data_Strage_3(I, J) = "R_Fsp_r :Cr"

  ' Tilting_os Stability_os
   I = I + 1:   Data_Strage_3(I, J) = "Tilting_os"      '"Ctr*:Sqrt(Ct2+Cr2)"
   I = I + 1:   Data_Strage_3(I, J) = "Stability_os"    '"ξ*:Sqrt(Ct2+Cr2)/Ros"

  ' Driving Torque      20180711
   I = I + 1:   Data_Strage_3(I, J) = "Torque_s"        ' =Torque_s = R_Fsp_r * Ro

  ' Difference pressure of Wrap tips
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(1)"        ' =DP_tp(1)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(2)"        ' =DP_tp(2)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(3)"        ' =DP_tp(3)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(4)"        ' =DP_tp(4)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(5)"        ' =DP_tp(5)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tp(6)"        ' =DP_tp(6)

   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(1)"        ' =DP_tg(1)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(2)"        ' =DP_tg(2)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(3)"        ' =DP_tg(3)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(4)"        ' =DP_tg(4)
   I = I + 1:   Data_Strage_3(I, J) = "DP_tg(5)"        ' =DP_tg(5)

  ' Load of os Eccentric bearing
   I = I + 1:   Data_Strage_3(I, J) = "Fsb_e"           ' =Fsb_e = Sqr(R_Fsp_r ^ 2 + R_Fsp_r ^ 2)



'--------------------------
'-- Data 部
'--------------------------

        I = 3     ' 代入する配列の開始行
        J = 1 + 1   ' 代入する配列の開始列

 '------ 1-20 [項目名]  A室側 外から１番目の接点

               Data_Strage_3(I, J) = N_rps               ' #1
   I = I + 1:   Data_Strage_3(I, J) = Hw                 ' #2
   I = I + 1:   Data_Strage_3(I, J) = P_discharge - atm_00               ' #3
   I = I + 1:   Data_Strage_3(I, J) = P_suction - atm_00               ' #4
   I = I + 1:   Data_Strage_3(I, J) = OS_seal                 ' #5
   I = I + 1:   Data_Strage_3(I, J) = Ro                 ' #6
   I = I + 1:   Data_Strage_3(I, J) = delta_ky * 180 / pi             ' #7
   I = I + 1:   Data_Strage_3(I, J) = alpha_ky * 180 / pi             ' #8
   I = I + 1:   Data_Strage_3(I, J) = myu_sb                 ' #9
   I = I + 1:   Data_Strage_3(I, J) = myu_ky                 ' #10
   I = I + 1:   Data_Strage_3(I, J) = myu_th                 ' #11
   I = I + 1:   Data_Strage_3(I, J) = h_pl                 ' #12
   I = I + 1:   Data_Strage_3(I, J) = OS_dia / 2               ' #13
   I = I + 1:   Data_Strage_3(I, J) = h_ky                 ' #14
   I = I + 1:   Data_Strage_3(I, J) = b_kos                 ' #15
   I = I + 1:   Data_Strage_3(I, J) = b_kmf                 ' #16
   I = I + 1:   Data_Strage_3(I, J) = Z_eb                 ' #17
   I = I + 1:   Data_Strage_3(I, J) = R_kmf                 ' #18
   I = I + 1:   Data_Strage_3(I, J) = R_kos                 ' #19
   I = I + 1:   Data_Strage_3(I, J) = Fs_r                 ' #20
   I = I + 1:   Data_Strage_3(I, J) = Index_I                 ' #21
   I = I + 1:   Data_Strage_3(I, J) = N_wrap_a(Index_I)                 ' #22
   I = I + 1:   Data_Strage_3(I, J) = N_wrap_b(Index_I)                 ' #23
   I = I + 1:   Data_Strage_3(I, J) = the_c(Index_I) * 180 / pi           ' #24
   I = I + 1:   Data_Strage_3(I, J) = x_e(Index_I) * (-1)               ' #25
   I = I + 1:   Data_Strage_3(I, J) = y_e(Index_I)                 ' #26
   I = I + 1:   Data_Strage_3(I, J) = Press_A(3, Index_I) - atm_00               ' #27
   I = I + 1:   Data_Strage_3(I, J) = Press_A(2, Index_I) - atm_00               ' #28
   I = I + 1:   Data_Strage_3(I, J) = Press_A(1, Index_I) - atm_00               ' #29
   I = I + 1:   Data_Strage_3(I, J) = Press_A(0, Index_I) - atm_00               ' #30
   I = I + 1:   Data_Strage_3(I, J) = Press_B(3, Index_I) - atm_00               ' #31
   I = I + 1:   Data_Strage_3(I, J) = Press_B(2, Index_I) - atm_00               ' #32
   I = I + 1:   Data_Strage_3(I, J) = Press_B(1, Index_I) - atm_00               ' #33
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(1, Index_I)                 ' #34
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(2, Index_I)                 ' #35
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(3, Index_I)                 ' #36
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(4, Index_I)                 ' #37
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(5, Index_I)                 ' #38
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(6, Index_I)                 ' #39
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(1, Index_I)                 ' #40
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(2, Index_I)                 ' #41
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(3, Index_I)                 ' #42
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(4, Index_I)                 ' #43
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(5, Index_I)                 ' #44
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(6, Index_I)                 ' #45
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(7, Index_I)                 ' #46
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(8, Index_I)                 ' #47
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(9, Index_I)                 ' #48
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(10, Index_I)                 ' #49
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(11, Index_I)                 ' #50
   I = I + 1:   Data_Strage_3(I, J) = P_groove - atm_00               ' #51
   I = I + 1:   Data_Strage_3(I, J) = P_back - atm_00               ' #52
   I = I + 1:   Data_Strage_3(I, J) = Fgc_r                 ' #53
   I = I + 1:   Data_Strage_3(I, J) = Fgc_t                 ' #54
   I = I + 1:   Data_Strage_3(I, J) = R_Fgc_t                 ' #55
   I = I + 1:   Data_Strage_3(I, J) = R_Fgc_r                 ' #56
   I = I + 1:   Data_Strage_3(I, J) = Fgc_t * R_Fgc_t               ' #57
   I = I + 1:   Data_Strage_3(I, J) = Fgc_r * R_Fgc_r               ' #58
   I = I + 1:   Data_Strage_3(I, J) = xg2_a(0, Index_I) * (-1)               ' #59
   I = I + 1:   Data_Strage_3(I, J) = yg2_a(0, Index_I)                 ' #60
   I = I + 1:   Data_Strage_3(I, J) = Sg2_a(0, Index_I)                 ' #61
   I = I + 1:   Data_Strage_3(I, J) = Pz_A(0, Index_I)                 ' #62
   I = I + 1:   Data_Strage_3(I, J) = Fz_A(0, Index_I)                 ' #63
   I = I + 1:   Data_Strage_3(I, J) = Mzx_A(0, Index_I) * (-1)               ' #64
   I = I + 1:   Data_Strage_3(I, J) = Mzy_A(0, Index_I)                 ' #65
   I = I + 1:   Data_Strage_3(I, J) = xg2_a(2, Index_I) * (-1)               ' #66
   I = I + 1:   Data_Strage_3(I, J) = yg2_a(2, Index_I)                 ' #67
   I = I + 1:   Data_Strage_3(I, J) = Sg2_a(2, Index_I)                 ' #68
   I = I + 1:   Data_Strage_3(I, J) = Pz_A(2, Index_I)                 ' #69
   I = I + 1:   Data_Strage_3(I, J) = Fz_A(2, Index_I)                 ' #70
   I = I + 1:   Data_Strage_3(I, J) = Mzx_A(2, Index_I) * (-1)               ' #71
   I = I + 1:   Data_Strage_3(I, J) = Mzy_A(2, Index_I)                 ' #72
   I = I + 1:   Data_Strage_3(I, J) = xg2_a(3, Index_I) * (-1)               ' #73
   I = I + 1:   Data_Strage_3(I, J) = yg2_a(3, Index_I)                 ' #74
   I = I + 1:   Data_Strage_3(I, J) = Sg2_a(3, Index_I)                 ' #75
   I = I + 1:   Data_Strage_3(I, J) = Pz_A(3, Index_I)                 ' #76
   I = I + 1:   Data_Strage_3(I, J) = Fz_A(3, Index_I)                 ' #77
   I = I + 1:   Data_Strage_3(I, J) = Mzx_A(3, Index_I) * (-1)               ' #78
   I = I + 1:   Data_Strage_3(I, J) = Mzy_A(3, Index_I)                 ' #79
   I = I + 1:   Data_Strage_3(I, J) = xg2_a(4, Index_I) * (-1)               ' #80
   I = I + 1:   Data_Strage_3(I, J) = yg2_a(4, Index_I)                 ' #81
   I = I + 1:   Data_Strage_3(I, J) = Sg2_a(4, Index_I)                 ' #82
   I = I + 1:   Data_Strage_3(I, J) = Pz_A(4, Index_I)                 ' #83
   I = I + 1:   Data_Strage_3(I, J) = Fz_A(4, Index_I)                 ' #84
   I = I + 1:   Data_Strage_3(I, J) = Mzx_A(4, Index_I)                 ' #85
   I = I + 1:   Data_Strage_3(I, J) = Mzy_A(4, Index_I)                 ' #86
   I = I + 1:   Data_Strage_3(I, J) = xg2_b(2, Index_I) * (-1)               ' #87
   I = I + 1:   Data_Strage_3(I, J) = yg2_b(2, Index_I)                 ' #88
   I = I + 1:   Data_Strage_3(I, J) = Sg2_b(2, Index_I)                 ' #89
   I = I + 1:   Data_Strage_3(I, J) = Pz_B(2, Index_I)                 ' #90
   I = I + 1:   Data_Strage_3(I, J) = Fz_B(2, Index_I)                 ' #91
   I = I + 1:   Data_Strage_3(I, J) = Mzx_B(2, Index_I) * (-1)               ' #92
   I = I + 1:   Data_Strage_3(I, J) = Mzy_B(2, Index_I)                 ' #93
   I = I + 1:   Data_Strage_3(I, J) = xg2_b(3, Index_I) * (-1)               ' #94
   I = I + 1:   Data_Strage_3(I, J) = yg2_b(3, Index_I)                 ' #95
   I = I + 1:   Data_Strage_3(I, J) = Sg2_b(3, Index_I)                 ' #96
   I = I + 1:   Data_Strage_3(I, J) = Pz_B(3, Index_I)                 ' #97
   I = I + 1:   Data_Strage_3(I, J) = Fz_B(3, Index_I)                 ' #98
   I = I + 1:   Data_Strage_3(I, J) = Mzx_B(3, Index_I) * (-1)               ' #99
   I = I + 1:   Data_Strage_3(I, J) = Mzy_B(3, Index_I)                 ' #100
   I = I + 1:   Data_Strage_3(I, J) = xg_m(6, Index_I) * (-1)               ' #101
   I = I + 1:   Data_Strage_3(I, J) = yg_m(6, Index_I)                 ' #102
   I = I + 1:   Data_Strage_3(I, J) = Sg_m(6, Index_I)                 ' #103
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(6, Index_I)                 ' #104
   I = I + 1:   Data_Strage_3(I, J) = Fz_m(6, Index_I)                 ' #105
   I = I + 1:   Data_Strage_3(I, J) = Mzx_m(6, Index_I)                 ' #106
   I = I + 1:   Data_Strage_3(I, J) = Mzy_m(6, Index_I)                 ' #107
   I = I + 1:   Data_Strage_3(I, J) = xg_m(5, Index_I) * (-1)               ' #108
   I = I + 1:   Data_Strage_3(I, J) = yg_m(5, Index_I)                 ' #109
   I = I + 1:   Data_Strage_3(I, J) = Sg_m(5, Index_I)                 ' #110
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(5, Index_I)                 ' #111
   I = I + 1:   Data_Strage_3(I, J) = Fz_m(5, Index_I)                 ' #112
   I = I + 1:   Data_Strage_3(I, J) = Mzx_m(5, Index_I) * (-1)               ' #113
   I = I + 1:   Data_Strage_3(I, J) = Mzy_m(5, Index_I)                 ' #114
   I = I + 1:   Data_Strage_3(I, J) = xg_m(4, Index_I) * (-1)               ' #115
   I = I + 1:   Data_Strage_3(I, J) = yg_m(4, Index_I)                 ' #116
   I = I + 1:   Data_Strage_3(I, J) = Sg_m(4, Index_I)                 ' #117
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(4, Index_I)                 ' #118
   I = I + 1:   Data_Strage_3(I, J) = Fz_m(4, Index_I)                 ' #119
   I = I + 1:   Data_Strage_3(I, J) = Mzx_m(4, Index_I) * (-1)               ' #120
   I = I + 1:   Data_Strage_3(I, J) = Mzy_m(4, Index_I)                 ' #121
   I = I + 1:   Data_Strage_3(I, J) = xg_m(3, Index_I) * (-1)               ' #122
   I = I + 1:   Data_Strage_3(I, J) = yg_m(3, Index_I)                 ' #123
   I = I + 1:   Data_Strage_3(I, J) = Sg_m(3, Index_I)                 ' #124
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(3, Index_I)                 ' #125
   I = I + 1:   Data_Strage_3(I, J) = Fz_m(3, Index_I)                 ' #126
   I = I + 1:   Data_Strage_3(I, J) = Mzx_m(3, Index_I) * (-1)               ' #127
   I = I + 1:   Data_Strage_3(I, J) = Mzy_m(3, Index_I)                 ' #128
   I = I + 1:   Data_Strage_3(I, J) = xg_m(2, Index_I) * (-1)               ' #129
   I = I + 1:   Data_Strage_3(I, J) = yg_m(2, Index_I)                 ' #130
   I = I + 1:   Data_Strage_3(I, J) = Sg_m(2, Index_I)                 ' #131
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(2, Index_I)                 ' #132
   I = I + 1:   Data_Strage_3(I, J) = Fz_m(2, Index_I)                 ' #133
   I = I + 1:   Data_Strage_3(I, J) = Mzx_m(2, Index_I) * (-1)               ' #134
   I = I + 1:   Data_Strage_3(I, J) = Mzy_m(2, Index_I)                 ' #135
   I = I + 1:   Data_Strage_3(I, J) = xg_m(1, Index_I) * (-1)               ' #136
   I = I + 1:   Data_Strage_3(I, J) = yg_m(1, Index_I)                 ' #137
   I = I + 1:   Data_Strage_3(I, J) = Sg_m(1, Index_I)                 ' #138
   I = I + 1:   Data_Strage_3(I, J) = Pz_m(1, Index_I)                 ' #139
   I = I + 1:   Data_Strage_3(I, J) = Fz_m(1, Index_I)                 ' #140
   I = I + 1:   Data_Strage_3(I, J) = Mzx_m(1, Index_I) * (-1)               ' #141
   I = I + 1:   Data_Strage_3(I, J) = Mzy_m(1, Index_I)                 ' #142
   I = I + 1:   Data_Strage_3(I, J) = xg_f(6, Index_I) * (-1)               ' #143
   I = I + 1:   Data_Strage_3(I, J) = yg_f(6, Index_I)                 ' #144
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(6, Index_I)                 ' #145
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(6, Index_I)                 ' #146
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(6, Index_I)                 ' #147
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(6, Index_I) * (-1)               ' #148
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(6, Index_I)                 ' #149
   I = I + 1:   Data_Strage_3(I, J) = xg_f(5, Index_I) * (-1)               ' #150
   I = I + 1:   Data_Strage_3(I, J) = yg_f(5, Index_I)                 ' #151
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(5, Index_I)                 ' #152
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(5, Index_I)                 ' #153
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(5, Index_I)                 ' #154
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(5, Index_I) * (-1)               ' #155
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(5, Index_I)                 ' #156
   I = I + 1:   Data_Strage_3(I, J) = xg_f(4, Index_I) * (-1)               ' #157
   I = I + 1:   Data_Strage_3(I, J) = yg_f(4, Index_I)                 ' #158
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(4, Index_I)                 ' #159
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(4, Index_I)                 ' #160
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(4, Index_I)                 ' #161
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(4, Index_I) * (-1)               ' #162
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(4, Index_I)                 ' #163
   I = I + 1:   Data_Strage_3(I, J) = xg_f(3, Index_I) * (-1)               ' #164
   I = I + 1:   Data_Strage_3(I, J) = yg_f(3, Index_I)                 ' #165
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(3, Index_I)                 ' #166
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(3, Index_I)                 ' #167
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(3, Index_I)                 ' #168
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(3, Index_I) * (-1)               ' #169
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(3, Index_I)                 ' #170
   I = I + 1:   Data_Strage_3(I, J) = xg_f(2, Index_I) * (-1)               ' #171
   I = I + 1:   Data_Strage_3(I, J) = yg_f(2, Index_I)                 ' #172
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(2, Index_I)                 ' #173
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(2, Index_I)                 ' #174
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(2, Index_I)                 ' #175
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(2, Index_I) * (-1)               ' #176
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(2, Index_I)                 ' #177
   I = I + 1:   Data_Strage_3(I, J) = xg_f(1, Index_I) * (-1)               ' #178
   I = I + 1:   Data_Strage_3(I, J) = yg_f(1, Index_I)                 ' #179
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(1, Index_I)                 ' #180
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(1, Index_I)                 ' #181
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(1, Index_I)                 ' #182
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(1, Index_I) * (-1)               ' #183
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(1, Index_I)                 ' #184
   I = I + 1:   Data_Strage_3(I, J) = xg_f(7, Index_I) * (-1)               ' #185
   I = I + 1:   Data_Strage_3(I, J) = yg_f(7, Index_I)                 ' #186
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(7, Index_I)                 ' #187
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(7, Index_I)                 ' #188
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(7, Index_I)                 ' #189
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(7, Index_I) * (-1)               ' #190
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(7, Index_I)                 ' #191
   I = I + 1:   Data_Strage_3(I, J) = xg_f(8, Index_I) * (-1)               ' #192
   I = I + 1:   Data_Strage_3(I, J) = yg_f(8, Index_I)                 ' #193
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(8, Index_I)                 ' #194
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(8, Index_I)                 ' #195
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(8, Index_I)                 ' #196
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(8, Index_I) * (-1)               ' #197
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(8, Index_I)                 ' #198
   I = I + 1:   Data_Strage_3(I, J) = xg_f(9, Index_I) * (-1)               ' #199
   I = I + 1:   Data_Strage_3(I, J) = yg_f(9, Index_I)                 ' #200
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(9, Index_I)                 ' #201
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(9, Index_I)                 ' #202
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(9, Index_I)                 ' #203
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(9, Index_I) * (-1)               ' #204
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(9, Index_I)                 ' #205
   I = I + 1:   Data_Strage_3(I, J) = xg_f(10, Index_I) * (-1)               ' #206
   I = I + 1:   Data_Strage_3(I, J) = yg_f(10, Index_I)                 ' #207
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(10, Index_I)                 ' #208
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(10, Index_I)                 ' #209
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(10, Index_I)                 ' #210
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(10, Index_I) * (-1)               ' #211
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(10, Index_I)                 ' #212
   I = I + 1:   Data_Strage_3(I, J) = xg_f(11, Index_I) * (-1)               ' #213
   I = I + 1:   Data_Strage_3(I, J) = yg_f(11, Index_I)                 ' #214
   I = I + 1:   Data_Strage_3(I, J) = Sg_f(11, Index_I)                 ' #215
   I = I + 1:   Data_Strage_3(I, J) = Pz_f(11, Index_I)                 ' #216
   I = I + 1:   Data_Strage_3(I, J) = Fz_f(11, Index_I)                 ' #217
   I = I + 1:   Data_Strage_3(I, J) = Mzx_f(11, Index_I)                 ' #218
   I = I + 1:   Data_Strage_3(I, J) = Mzy_f(11, Index_I)                 ' #219
   I = I + 1:   Data_Strage_3(I, J) = xg2_b(7, Index_I) * (-1)               ' #220
   I = I + 1:   Data_Strage_3(I, J) = yg2_b(7, Index_I)                 ' #221
   I = I + 1:   Data_Strage_3(I, J) = Sg2_b(7, Index_I)                 ' #222
   I = I + 1:   Data_Strage_3(I, J) = Pz_B(7, Index_I)                 ' #223
   I = I + 1:   Data_Strage_3(I, J) = Fz_B(7, Index_I)                 ' #224
   I = I + 1:   Data_Strage_3(I, J) = Mzx_B(7, Index_I) * (-1)               ' #225
   I = I + 1:   Data_Strage_3(I, J) = Mzy_B(7, Index_I)                 ' #226
   I = I + 1:   Data_Strage_3(I, J) = Fz_Za(Index_I)                 ' #227
   I = I + 1:   Data_Strage_3(I, J) = Mzx_Za(Index_I) * (-1)               ' #228
   I = I + 1:   Data_Strage_3(I, J) = Mzy_Za(Index_I)                 ' #229
   I = I + 1:   Data_Strage_3(I, J) = x_zforce_a(Index_I) * (-1)               ' #230
   I = I + 1:   Data_Strage_3(I, J) = y_zforce_a(Index_I)                 ' #231
   I = I + 1:   Data_Strage_3(I, J) = (x_zforce_a(Index_I) - x_e(Index_I)) * (-1)             ' #232
   I = I + 1:   Data_Strage_3(I, J) = y_zforce_a(Index_I) - y_e(Index_I)               ' #233
   I = I + 1:   Data_Strage_3(I, J) = r_zforce_a(Index_I)                 ' #234
   I = I + 1:   Data_Strage_3(I, J) = t_zforce_a(Index_I) * (-1)                 ' #235
   I = I + 1:   Data_Strage_3(I, J) = xg2_b(8, Index_I)                 ' #236
   I = I + 1:   Data_Strage_3(I, J) = yg2_b(8, Index_I) * (-1)               ' #237
   I = I + 1:   Data_Strage_3(I, J) = Sg2_b(8, Index_I)                 ' #238
   I = I + 1:   Data_Strage_3(I, J) = Pz_B(8, Index_I)                 ' #239
   I = I + 1:   Data_Strage_3(I, J) = Fz_B(8, Index_I)                 ' #240
   I = I + 1:   Data_Strage_3(I, J) = Mzx_B(8, Index_I) * (-1)               ' #241
   I = I + 1:   Data_Strage_3(I, J) = Mzy_B(8, Index_I)                 ' #242
   I = I + 1:   Data_Strage_3(I, J) = xg2_b(9, Index_I)                 ' #243
   I = I + 1:   Data_Strage_3(I, J) = yg2_b(9, Index_I)                 ' #244
   I = I + 1:   Data_Strage_3(I, J) = Sg2_b(9, Index_I)                 ' #245
   I = I + 1:   Data_Strage_3(I, J) = Pz_B(9, Index_I)                 ' #246
   I = I + 1:   Data_Strage_3(I, J) = Fz_B(9, Index_I)                 ' #247
   I = I + 1:   Data_Strage_3(I, J) = Mzx_B(9, Index_I)                 ' #248
   I = I + 1:   Data_Strage_3(I, J) = Mzy_B(9, Index_I)                 ' #249
   I = I + 1:   Data_Strage_3(I, J) = 0                 ' #250
   I = I + 1:   Data_Strage_3(I, J) = 0                 ' #251
   I = I + 1:   Data_Strage_3(I, J) = 0                 ' #252
   I = I + 1:   Data_Strage_3(I, J) = 0                 ' #253
   I = I + 1:   Data_Strage_3(I, J) = 0                 ' #254
   I = I + 1:   Data_Strage_3(I, J) = 0                 ' #255
   I = I + 1:   Data_Strage_3(I, J) = 0                 ' #256
   I = I + 1:   Data_Strage_3(I, J) = Fz_Zb(Index_I)                 ' #257
   I = I + 1:   Data_Strage_3(I, J) = Mzx_Zb(Index_I) * (-1)               ' #258
   I = I + 1:   Data_Strage_3(I, J) = Mzy_Zb(Index_I)                 ' #259
   I = I + 1:   Data_Strage_3(I, J) = x_zforce_b(Index_I) * (-1)               ' #260
   I = I + 1:   Data_Strage_3(I, J) = y_zforce_b(Index_I)                 ' #261
   I = I + 1:   Data_Strage_3(I, J) = (x_zforce_b(Index_I) - x_e(Index_I)) * (-1)             ' #262
   I = I + 1:   Data_Strage_3(I, J) = y_zforce_b(Index_I) - y_e(Index_I)               ' #263
   I = I + 1:   Data_Strage_3(I, J) = r_zforce_b(Index_I)                 ' #264
   I = I + 1:   Data_Strage_3(I, J) = t_zforce_b(Index_I)                 ' #265
   I = I + 1:   Data_Strage_3(I, J) = Fz_sp(Index_I)                 ' #266


   I = I + 1:   Data_Strage_3(I, J) = R_x_osw(1) * (-1)                 ' #267
   I = I + 1:   Data_Strage_3(I, J) = R_y_osw(1)                 ' #268
   I = I + 1:   Data_Strage_3(I, J) = R_z_osw(1)                 ' #269
   I = I + 1:   Data_Strage_3(I, J) = V_osw(1)                 ' #270
   I = I + 1:   Data_Strage_3(I, J) = R_x_osw(2) * (-1)                 ' #271
   I = I + 1:   Data_Strage_3(I, J) = R_y_osw(2)                 ' #272
   I = I + 1:   Data_Strage_3(I, J) = R_z_osw(2)                 ' #273
   I = I + 1:   Data_Strage_3(I, J) = V_osw(2)                 ' #274
   I = I + 1:   Data_Strage_3(I, J) = R_x_osw(4) * (-1)                 ' #275
   I = I + 1:   Data_Strage_3(I, J) = R_y_osw(4)                 ' #276
   I = I + 1:   Data_Strage_3(I, J) = R_z_osw(4)                 ' #277
   I = I + 1:   Data_Strage_3(I, J) = V_osw(4)                 ' #278

   I = I + 1:   Data_Strage_3(I, J) = R_x_osw(0) * (-1)               ' #279
   I = I + 1:   Data_Strage_3(I, J) = R_y_osw(0)                 ' #280
   I = I + 1:   Data_Strage_3(I, J) = R_z_osw(0)                 ' #281
   I = I + 1:   Data_Strage_3(I, J) = R_Fmg_r                 ' #282
   I = I + 1:   Data_Strage_3(I, J) = R_Fmg_t                 ' #283
   I = I + 1:   Data_Strage_3(I, J) = vol_os                 ' #284
   I = I + 1:   Data_Strage_3(I, J) = m_os                 ' #285
   I = I + 1:   Data_Strage_3(I, J) = F_mg                 ' #286
   I = I + 1:   Data_Strage_3(I, J) = Fmc_r                 ' #287


   I = I + 1:   Data_Strage_3(I, J) = Ros_F1_oy                 ' #288
   I = I + 1:   Data_Strage_3(I, J) = Ros_F2_oy                 ' #289
   I = I + 1:   Data_Strage_3(I, J) = Ros_F3_ox                 ' #290
   I = I + 1:   Data_Strage_3(I, J) = Ros_F4_ox                 ' #291
   I = I + 1:   Data_Strage_3(I, J) = R_x_or(1)                 ' #292  '[0719] * (-1)
   I = I + 1:   Data_Strage_3(I, J) = R_y_or(1) * (-1)          ' #293
   I = I + 1:   Data_Strage_3(I, J) = R_z_or(1)                 ' #294
   I = I + 1:   Data_Strage_3(I, J) = V_or(1)                 ' #295
   I = I + 1:   Data_Strage_3(I, J) = R_x_or(2)                 ' #296  '[0719]* (-1)
   I = I + 1:   Data_Strage_3(I, J) = R_y_or(2) * (-1)          ' #297
   I = I + 1:   Data_Strage_3(I, J) = R_z_or(2)                 ' #298
   I = I + 1:   Data_Strage_3(I, J) = V_or(2)                 ' #299
   I = I + 1:   Data_Strage_3(I, J) = R_x_or(3)                 ' #300  '[0719]* (-1)
   I = I + 1:   Data_Strage_3(I, J) = R_y_or(3) * (-1)          ' #301
   I = I + 1:   Data_Strage_3(I, J) = R_z_or(3)                 ' #302
   I = I + 1:   Data_Strage_3(I, J) = V_or(3)                 ' #303

   I = I + 1:   Data_Strage_3(I, J) = R_x_or(0)                 ' #304  '[0719]* (-1)
   I = I + 1:   Data_Strage_3(I, J) = R_y_or(0) * (-1)          ' #305
   I = I + 1:   Data_Strage_3(I, J) = R_z_or(0)                 ' #306
   I = I + 1:   Data_Strage_3(I, J) = vol_or    'V_or(3)      ' #307
   I = I + 1:   Data_Strage_3(I, J) = m_or                      ' #308
   I = I + 1:   Data_Strage_3(I, J) = Fc_or                     ' #309   '[0719]*(-1)

  ' Matrix
   I = I + 1:   Data_Strage_3(I, J) = Matrix_X(1)   ' =Fk_1     "F1"
   I = I + 1:   Data_Strage_3(I, J) = Matrix_X(2)   ' =Fk_2     "F2"
   I = I + 1:   Data_Strage_3(I, J) = Matrix_X(3)   ' =Fk_3     "F3"
   I = I + 1:   Data_Strage_3(I, J) = Matrix_X(4)   ' =Fk_4     "F4"
   I = I + 1:   Data_Strage_3(I, J) = Matrix_X(5)   ' =Fsb_t    "Fsbt"
   I = I + 1:   Data_Strage_3(I, J) = Matrix_X(6)   ' =Fsb_r    "Fsbr"
   I = I + 1:   Data_Strage_3(I, J) = Matrix_X(7)   ' =R_Fsp_t  "Ct"
   I = I + 1:   Data_Strage_3(I, J) = Matrix_X(8)   ' =R_Fsp_r  "Cr" [317]

  ' Tilting_os Stability_os   [318]
   I = I + 1:   Data_Strage_3(I, J) _
                  = Sqr(Matrix_X(7) ^ 2 + Matrix_X(8) ^ 2)                ' "Ctr*:Sqrt(Ct2+Cr2)"
   I = I + 1:   Data_Strage_3(I, J) _
                  = Sqr(Matrix_X(7) ^ 2 + Matrix_X(8) ^ 2) * 2 / (OS_dia) ' "ξ*:Sqrt(Ct2+Cr2)/Ros" [319]

  ' Driving Torque      [320]   20180711
   I = I + 1:   Data_Strage_3(I, J) = Torque_s        ' =Torque_s = R_Fsp_r * Ro

  ' Difference pressure of Wrap tips
   I = I + 1:   Data_Strage_3(I, J) = DP_tm(1, Index_I)        ' =DP_tp(1)    [321]
   I = I + 1:   Data_Strage_3(I, J) = DP_tm(2, Index_I)        ' =DP_tp(2)
   I = I + 1:   Data_Strage_3(I, J) = DP_tm(3, Index_I)        ' =DP_tp(3)
   I = I + 1:   Data_Strage_3(I, J) = DP_tm(4, Index_I)        ' =DP_tp(4)
   I = I + 1:   Data_Strage_3(I, J) = DP_tm(5, Index_I)        ' =DP_tp(5)
   I = I + 1:   Data_Strage_3(I, J) = DP_tm(6, Index_I)        ' =DP_tp(6)

   I = I + 1:   Data_Strage_3(I, J) = DP_tf(1, Index_I)        ' =DP_tg(1)    [327]
   I = I + 1:   Data_Strage_3(I, J) = DP_tf(2, Index_I)        ' =DP_tg(2)
   I = I + 1:   Data_Strage_3(I, J) = DP_tf(3, Index_I)        ' =DP_tg(3)
   I = I + 1:   Data_Strage_3(I, J) = DP_tf(4, Index_I)        ' =DP_tg(4)
   I = I + 1:   Data_Strage_3(I, J) = DP_tf(5, Index_I)        ' =DP_tg(5)

  ' Load of os Eccentric bearing      [332]
   I = I + 1:   Data_Strage_3(I, J) = Fsb_e           ' =Fsb_e = Sqr(R_Fsp_r ^ 2 + R_Fsp_r ^ 2)


'--------------------------
'-- Data 一括貼付
'--------------------------


    Sheets(DataSheetName_3).Select
'    Sheets(DataSheetName_3).Cells.ClearContents                    '：Sheet全Cellの数式、文字列をクリア
    Sheets(DataSheetName_3).Range("H1:I999").ClearContents        '：指定Cellの数式、文字列をクリア

    I1 = 1  'Range("A1").row       ' 貼付先の先頭セルの、列と行 Cells(i1, j1)
    J1 = 8  'Range("H1").Column
            ' 参考) Cells(2, 1).Select = Range("A2").Select　　'Cells(8,3) = Range("H3")

        With Sheets(DataSheetName_3)
            .Range(Cells(I1, J1), Cells(I1 + dw_n3 - 1, J1 + dw_c - 1)).Value _
                = Data_Strage_3
'                = WorksheetFunction.Transpose(Data_Strage_3)
        End With

        With Sheets(DataSheetName_3)
            .Cells(1, J1).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

'            .Cells(1, 3).Value = Format(Now(), "yyyy/MM/DD")       '　"Date
'            .Cells(1, 4).Value = Format(Now(), "HH:mm:ss")         '　"Date
        End With

'    Sheets(DataSheetName_3).Range("A117:CM999").ClearContents        '：指定Cellの数式、文字列をクリア

End Sub





'======================================================== 【M4-2】
'  結果を、配列に保存
'　　　　配列を、直接Cellへ格納
'
'　　　　※注）配列とCellの行と列は逆
'========================================================

Public Sub Data_Strage_to_array_4()

Dim I As Long, J As Long
Dim I1 As Long, J1 As Long
Dim ii_tmp As Long, jj_temp As Long

'【Cellに書き出し】

'--------------------------
'  Data Title 部
'--------------------------
    J = 2     '配列の開始列

    dw_c = 0            ' =0　配列の列数
    dw_c = dw_c + 12 + 36  ' =47　配列の列数


    ReDim Data_Strage(dw_c, dw_n + 3)

        I = 1:       Data_Strage(I, J) = "Index No."           '[1]
        I = I + 1:   Data_Strage(I, J) = "Phi2[deg]"           ' "Phi2_[deg]"   20171031
        I = I + 1:   Data_Strage(I, J) = "Phi1[deg]"           ' "Phi1_[deg]"   20171031
        I = I + 1:   Data_Strage(I, J) = "Phi2"                ' "Phi2_"        20171031
        I = I + 1:   Data_Strage(I, J) = "Phi1"                ' "Phi1_"        20171031
        I = I + 1:   Data_Strage(I, J) = "Theta"       ' 軸回転角　吸込=0 deg ※The()はφ2 基準
        I = I + 1:   Data_Strage(I, J) = "Theta_c [deg]"       ' 軸回転角　吸込=0 deg ※The()はφ2 基準

        I = I + 1:   Data_Strage(I, J) = "Phi_c_fi(3, i)"
        I = I + 1:   Data_Strage(I, J) = "Phi_c_fi(4, i)"
        I = I + 1:   Data_Strage(I, J) = "Phi_c_fi(5, i)"
        I = I + 1:   Data_Strage(I, J) = "Phi_c_fi(6, i)"
        I = I + 1:   Data_Strage(I, J) = "Phi_c_fi(7, i)"    '[12]

      For ii_tmp = 0 To 2
        I = I + 1:            Data_Strage(I, J) = "A_xfi_c(" & ii_tmp & ",j)"      '[13] [25]
        I = I + 1:            Data_Strage(I, J) = "A_yfi_c(" & ii_tmp & ",j)"
        I = I + 1:            Data_Strage(I, J) = "A_xmo_c(" & ii_tmp & ",j)"
        I = I + 1:            Data_Strage(I, J) = "A_ymo_c(" & ii_tmp & ",j)"

        I = I + 1:            Data_Strage(I, J) = "B_xmi_c(" & ii_tmp & ",j)"
        I = I + 1:            Data_Strage(I, J) = "B_ymi_c(" & ii_tmp & ",j)"
        I = I + 1:            Data_Strage(I, J) = "B_xfo_c(" & ii_tmp & ",j)"
        I = I + 1:            Data_Strage(I, J) = "B_yfo_c(" & ii_tmp & ",j)"

        I = I + 1:            Data_Strage(I, J) = "A_RR_fi_c(" & ii_tmp & ",j)"
        I = I + 1:            Data_Strage(I, J) = "A_RR_mo_c(" & ii_tmp & ",j)"
        I = I + 1:            Data_Strage(I, J) = "B_RR_mi_c(" & ii_tmp & ",j)"
        I = I + 1:            Data_Strage(I, J) = "B_RR_fo_c(" & ii_tmp & ",j)"    '[24] [36] [48]

      Next ii_tmp


'        i = i + 1:   Data_Strage(i, j) = "xg_f(0, j)"       '[10]
'        i = i + 1:   Data_Strage(i, j) = "yg_f(0, j)"
'        i = i + 1:   Data_Strage(i, j) = "Sg_f(0, j)"
'
'        i = i + 1:   Data_Strage(i, j) = "xg_m(0, j)"       '[13]
'        i = i + 1:   Data_Strage(i, j) = "yg_m(0, j)"
'        i = i + 1:   Data_Strage(i, j) = "Sg_m(0, j)"
'
'        i = i + 1:   Data_Strage(i, j) = "xg2_a(0, j)"      '[16]
'        i = i + 1:   Data_Strage(i, j) = "yg2_a(0, j)"
'        i = i + 1:   Data_Strage(i, j) = "Sg2_a(0, j)"      '[18]


'--------------------------
'  Data 部
'--------------------------

    For J = 0 To dw_n    ' the_c(0) to the_c(end= 366)

        I = 1:            Data_Strage(I, J + 3) = J                                '[1] "Index No."       20171031
        I = I + 1:        Data_Strage(I, J + 3) = phi2(J) * 180 / pi               ' "Phi2_[deg]"     20171031
        I = I + 1:        Data_Strage(I, J + 3) = phi1(J) * 180 / pi               ' "Phi2_[deg]"     20171031
        I = I + 1:        Data_Strage(I, J + 3) = phi2(J)                          ' "Phi2_"          20171031
        I = I + 1:        Data_Strage(I, J + 3) = phi1(J)                          ' "Phi1_"          20171031
        I = I + 1:        Data_Strage(I, J + 3) = the(J)
        I = I + 1:        Data_Strage(I, J + 3) = (the(0) - the(J)) * 180 / pi     ' the(0) - the(j)   20171031

        I = I + 1:        Data_Strage(I, J + 3) = Phi_c_fi(3, J)
        I = I + 1:        Data_Strage(I, J + 3) = Phi_c_fi(4, J)
        I = I + 1:        Data_Strage(I, J + 3) = Phi_c_fi(5, J)
        I = I + 1:        Data_Strage(I, J + 3) = Phi_c_fi(6, J)
        I = I + 1:        Data_Strage(I, J + 3) = Phi_c_fi(7, J)


            For ii_tmp = 0 To 2
               ' A chamber Wrap
              I = I + 1:        Data_Strage(I, J + 3) = xfi_c(ii_tmp, J)     '[13] [25]
              I = I + 1:        Data_Strage(I, J + 3) = yfi_c(ii_tmp, J)
              I = I + 1:        Data_Strage(I, J + 3) = xmo_c(ii_tmp, J)
              I = I + 1:        Data_Strage(I, J + 3) = ymo_c(ii_tmp, J)
               ' B chamber Wrap
              I = I + 1:        Data_Strage(I, J + 3) = xmi_c(ii_tmp, J)
              I = I + 1:        Data_Strage(I, J + 3) = ymi_c(ii_tmp, J)
              I = I + 1:        Data_Strage(I, J + 3) = xfo_c(ii_tmp, J)
              I = I + 1:        Data_Strage(I, J + 3) = yfo_c(ii_tmp, J)
               ' A chamber Wrap
              I = I + 1:        Data_Strage(I, J + 3) = RR_fi_c(ii_tmp, J)
              I = I + 1:        Data_Strage(I, J + 3) = RR_mo_c(ii_tmp, J)
               ' B chamber Wrap
              I = I + 1:        Data_Strage(I, J + 3) = RR_mi_c(ii_tmp, J)
              I = I + 1:        Data_Strage(I, J + 3) = RR_fo_c(ii_tmp, J)   '[24] [36] [48]

            Next ii_tmp


'        i = i + 1:         Data_Strage(i, j + 3) = xg_f(0, j)
'        i = i + 1:         Data_Strage(i, j + 3) = yg_f(0, j)
'        i = i + 1:         Data_Strage(i, j + 3) = Sg_f(0, j)
'
'        i = i + 1:         Data_Strage(i, j + 3) = xg_m(0, j)
'        i = i + 1:         Data_Strage(i, j + 3) = yg_m(0, j)
'        i = i + 1:         Data_Strage(i, j + 3) = Sg_m(0, j)
'
'        i = i + 1:         Data_Strage(i, j + 3) = xg2_a(0, j)
'        i = i + 1:         Data_Strage(i, j + 3) = yg2_a(0, j)
'        i = i + 1:         Data_Strage(i, j + 3) = Sg2_a(0, j)


'      If j > dw_n_PI(2) Then             ' the_c(0) to the_c(2PI) index( 0 to 183)
'         GoTo label_Data_Strage_end
'      End If
'
'        i = i + 1:            Data_Strage(i, j + 3) = LLt_A(0, j)
'        i = i + 1:            Data_Strage(i, j + 3) = LLt_A(1, j)
'        i = i + 1:            Data_Strage(i, j + 3) = LLt_A(2, j)
'        i = i + 1:            Data_Strage(i, j + 3) = Lt_D(j)         '※
'
'     ■■Label point
'     label_Data_Strage_end:

    Next J


'--------------------------
'  Data 一括貼付
'--------------------------

    Sheets(DataSheetName).Select
    Sheets(DataSheetName).Range("J4:CM999").ClearContents        '：指定Cellの数式、文字列をクリア

    I1 = 1                 ' 貼付先の先頭セルの、行と列 (i1, j1)
    J1 = 9
        With Sheets(DataSheetName)
            .Range(Cells(I1, J1), Cells(dw_n + 3 + I1, dw_c + J1)).Value _
                = WorksheetFunction.Transpose(Data_Strage)
        End With

        With Sheets(DataSheetName)
            .Cells(I1, J1 + 1).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

        End With

End Sub

'======================================================== 【    】
'  結果を、配列に保存
'　　　　配列を、直接Cellへ格納     ※重心結果
'
'　　　　※注）配列とCellの行と列は逆
'========================================================

Public Sub Data_Strage_to_array_3()

Dim I As Long, J As Long
Dim ii As Long, jj As Long

Dim I1 As Long, J1 As Long
Dim dw_n3 As Long

'【Cellに書き出し】

'--------------------------
'-- Data Title 部
'--------------------------

    dw_c = 3 + 10 + (3 * 11) + (3 * 7) + (3 * 8) + (3 * 8) ' 　配列の列数
    dw_n3 = dw_c           'dw_n3 = dw_n + 3

    ReDim Data_Strage_3(dw_c, dw_c)
'    ReDim Data_Strage_3(dw_c, dw_n3)       ' Data_Strage_3(dw_c, dw_n + 3)

        J = 2     '配列の開始列　2
        I = 1     '配列の開始行

 '------ 1-12 [項目名]  A室側 外から１番目の接点

                     Data_Strage_3(I, J) = "Index No."           '[1]
        I = I + 1:   Data_Strage_3(I, J) = "Phi2[deg]"           ' "Phi2_[deg]"   20171031
        I = I + 1:   Data_Strage_3(I, J) = "Theta_c [deg]"       ' 軸回転角　吸込=0 deg ※The()はφ2 基準

   ' OS Wrap　:tp1 ～ tp6
      For jj = 0 To 6
        I = I + 1:   Data_Strage_3(I, J) = "Sg_ｍ(" & jj & ",j)"
        I = I + 1:   Data_Strage_3(I, J) = "xg_ｍ(" & jj & ",j)"
        I = I + 1:   Data_Strage_3(I, J) = "yg_ｍ(" & jj & ",j)"
      Next jj

   ' FS Wrap　:tg1 ～ tg4 /  外周部 tg5 ～ tg10
      For jj = 0 To 10
        I = I + 1:   Data_Strage_3(I, J) = "Sg_f(" & jj & ",j)"
        I = I + 1:   Data_Strage_3(I, J) = "xg_f(" & jj & ",j)"
        I = I + 1:   Data_Strage_3(I, J) = "yg_f(" & jj & ",j)"
      Next jj

   ' A Chanmber 図心 :  Pp(2),Pp(3), Suction1,2,3,[7]all, / [2]Pp(2) as Sg2_a(2,j)
      For jj = 0 To 7
        I = I + 1:   Data_Strage_3(I, J) = "Sg2_a(" & jj & ",j)"
        I = I + 1:   Data_Strage_3(I, J) = "xg2_a(" & jj & ",j)"
        I = I + 1:   Data_Strage_3(I, J) = "yg2_a(" & jj & ",j)"
      Next jj

   ' B Chanmber 図心 :  Pg(2),Pg(3)(+inlet), [7]PW(oil groove), [8]ML(OS Plate)
      For jj = 1 To 10
        I = I + 1:   Data_Strage_3(I, J) = "Sg2_b(" & jj & ",j)"
        I = I + 1:   Data_Strage_3(I, J) = "xg2_b(" & jj & ",j)"
        I = I + 1:   Data_Strage_3(I, J) = "yg2_b(" & jj & ",j)"
      Next jj

'--------------------------
'-- Data 部
'--------------------------

    For J = 0 To 0  'dw_n    ' the_c(0) to the_c(end= 366)

        I = 1:            Data_Strage_3(I, J + 3) = J
        I = I + 1:        Data_Strage_3(I, J + 3) = phi2(J) * 180 / pi
        I = I + 1:        Data_Strage_3(I, J + 3) = (the(0) - the(J)) * 180 / pi

            '   i = i + 1:        Data_Strage(i, j + 3) = xg_a(jj, j)
            '   i = i + 1:        Data_Strage(i, j + 3) = yg_a(jj, j)
            '   i = i + 1:        Data_Strage(i, j + 3) = xg_b(jj, j)
            '   i = i + 1:        Data_Strage(i, j + 3) = yg_b(jj, j)

   ' OS Wrap　:tp1 ～ tp6
      For jj = 0 To 6
         I = I + 1:        Data_Strage_3(I, J + 3) = Sg_m(jj, J)   ' OS Wrap　:tp1 ～ tp6
         I = I + 1:        Data_Strage_3(I, J + 3) = xg_m(jj, J)   ' OS Wrap 図心x :tp1 ～ tp6
         I = I + 1:        Data_Strage_3(I, J + 3) = yg_m(jj, J)
      Next jj

   ' FS Wrap　:tg1 ～ tg4 /  外周部 tg5 ～ tg10
      For jj = 0 To 10
         I = I + 1:        Data_Strage_3(I, J + 3) = Sg_f(jj, J)   ' FS Wrap　:tg1 ～ tg10
         I = I + 1:        Data_Strage_3(I, J + 3) = xg_f(jj, J)   ' FS Wrap 図心x
         I = I + 1:        Data_Strage_3(I, J + 3) = yg_f(jj, J)
      Next jj

   ' A Chanmber 図心 :  Pp(2),Pp(3), Suction1,2,3,all, / Pp(2) --> Sg2_a(2,j)
      For jj = 0 To 7
        I = I + 1:        Data_Strage_3(I, J + 3) = Sg2_a(jj, J)  ' A Chanmber
        I = I + 1:        Data_Strage_3(I, J + 3) = xg2_a(jj, J)  ' A Chanmber 図心x
        I = I + 1:        Data_Strage_3(I, J + 3) = yg2_a(jj, J)
       Next jj

   ' B Chanmber 図心 :  Pg(2),Pg(3)(+inlet), [7]PW(oil groove), [8]ML(OS Plate)
      For jj = 1 To 10
        I = I + 1:        Data_Strage_3(I, J + 3) = Sg2_b(jj, J)  ' B Chanmber
        I = I + 1:        Data_Strage_3(I, J + 3) = xg2_b(jj, J)  ' B Chanmber 図心x
        I = I + 1:        Data_Strage_3(I, J + 3) = yg2_b(jj, J)
      Next jj


    Next J

'--------------------------
'-- Data 一括貼付
'--------------------------

    Sheets(DataSheetName_3).Select
'    Sheets(DataSheetName_3).Cells.ClearContents                    '：Sheet全Cellの数式、文字列をクリア
    Sheets(DataSheetName_3).Range("C4:CM999").ClearContents        '：指定Cellの数式、文字列をクリア

    I1 = 1                 ' 貼付先の先頭セルの、行と列 (i1, j1)
    J1 = 1
        With Sheets(DataSheetName_3)
            .Range(Cells(I1, J1), Cells(dw_n3 + 3 + I1, dw_c + J1)).Value _
                = Data_Strage_3
'                = WorksheetFunction.Transpose(Data_Strage_3)
        End With

        With Sheets(DataSheetName_3)
            .Cells(1, 2).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

'            .Cells(1, 3).Value = Format(Now(), "yyyy/MM/DD")       '　"Date
'            .Cells(1, 4).Value = Format(Now(), "HH:mm:ss")         '　"Date
        End With

    Sheets(DataSheetName_3).Range("A117:CM999").ClearContents        '：指定Cellの数式、文字列をクリア

End Sub



'======================================================== 【M4-2】
'  結果を、配列に保存
'　　　　配列を、直接Cellへ格納
'
'　　　　※注）配列とCellの行と列は逆
'========================================================

Public Sub Data_Strage_to_array()

Dim I As Long, J As Long
Dim I1 As Long, J1 As Long
Dim tmp_int As Long

'【Cellに書き出し】

'-- Data Title 部
'
        J = 2     '配列の開始列　2
        I = 1     '配列の開始行

    dw_c = J + 29 + 16 + 16 + 5  ' =68　配列の列数
    dw_c = dw_c + 12 + 12      ' =92　配列の列数  Ft,Fr
    dw_c = dw_c + 4            ' =96　配列の列数  Ft_D, Lt_D, Fr_D, Lr_D
    dw_c = dw_c + 15 + 15      ' =126　配列の列数  Lmt_A,B,D, Mmt_A,B,D, Lmr_D, Mmr_D
'    dw_c = dw_c + 12           ' =138   xg_a,yg_a
'    dw_c = dw_c + 8           ' =146   xg_f,yg_f
'    dw_c = dw_c + 8           ' =154   xfi2,yfi2 /  xmi2,ymi2
'    dw_c = dw_c + 4           ' =158  xfi3,yfi3 /
'    dw_c = dw_c + 6 + 2 + 4       ' =166  Lr_A() *6 , sin()*2+4/   20180702
'    dw_c = dw_c + 6 + 2 + 4 + 20     ' =186
    dw_c = 138 + 12

    ReDim Data_Strage(dw_c, dw_n + 3)

                     Data_Strage(I, J) = "Index No."           '[1]
        I = I + 1:   Data_Strage(I, J) = "Phi2[deg]"           ' "Phi2_[deg]"   20171031
        I = I + 1:   Data_Strage(I, J) = "Phi2"                ' "Phi2_"        20171031
        I = I + 1:   Data_Strage(I, J) = "Phi1"                ' "Phi1_"        20171031
        I = I + 1:   Data_Strage(I, J) = "Theta"       ' 軸回転角　吸込=0 deg ※The()はφ2 基準

        I = I + 1:   Data_Strage(I, J) = "Theta_c [deg]"       ' 軸回転角　吸込=0 deg ※The()はφ2 基準
        I = I + 1:   Data_Strage(I, J) = "V_a [cc]"
        I = I + 1:   Data_Strage(I, J) = "V_b [cc]"
        I = I + 1:   Data_Strage(I, J) = "V_d [cc]"

        I = I + 1:   Data_Strage(I, J) = "P_A(1,i)[MPa]"    '[10]
        I = I + 1:   Data_Strage(I, J) = "P_B(1,i)[MPa]"
        I = I + 1:   Data_Strage(I, J) = "P_A(2,i)[MPa]"
        I = I + 1:   Data_Strage(I, J) = "P_B(2,i)[MPa]"    '
        I = I + 1:   Data_Strage(I, J) = "P_D(i)[MPa]"    '="P_A(0,i)[MPa]"  '[ppi-Pg(1)] Discharge chamber Pressure

 '------ 15 [項目名]  A室側 外から１番目の接点　[項目名]
        I = I + 1:            Data_Strage(I, J) = "A_xfi_c(0,i)"     '[15]
        I = I + 1:            Data_Strage(I, J) = "A_yfi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "A_xmo_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "A_ymo_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "A_RR_fi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "A_RR_mo_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "A_beta_fi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "A_del_fi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "A_del_mo_c(0,i)"  '[23]

'------ 24 [項目名]  B室側 外から１番目の接点
        I = I + 1:            Data_Strage(I, J) = "B_xmi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "B_ymi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "B_xfo_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "B_yfo_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "B_RR_mi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "A_RR_fo_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "B_beta_mi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "B_del_mi_c(0,i)"
        I = I + 1:            Data_Strage(I, J) = "B_del_fo_c(0,i)"  '[32]

 '------ 33 [項目名]  A室側 外から２番目の接点
        I = I + 1:            Data_Strage(I, J) = "A_xfi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "A_yfi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "A_xmo_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "A_ymo_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "A_RR_fi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "A_RR_mo_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "A_beta_fi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "A_del_fi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "A_del_mo_c(1,i)"  '[41]

 '------ 42 [項目名]  B室側 外から２番目の接点
        I = I + 1:            Data_Strage(I, J) = "B_xmi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "B_ymi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "B_xfo_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "B_yfo_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "B_RR_mi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "B_RR_fo_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "B_beta_mi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "B_del_mi_c(1,i)"
        I = I + 1:            Data_Strage(I, J) = "B_del_fo_c(1,i)"  '[50]

 '------ 51 [項目名]  A室側 外から３番目の接点
        I = I + 1:            Data_Strage(I, J) = "A_xfi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "A_yfi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "A_xmo_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "A_ymo_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "A_RR_fi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "A_RR_mo_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "A_beta_fi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "A_del_fi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "A_del_mo_c(2,i)"  '[59]

' '------ 60 [項目名]  B室側 外から３番目の接点
        I = I + 1:            Data_Strage(I, J) = "B_xmi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "B_ymi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "B_xfo_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "B_yfo_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "B_RR_mi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "B_RR_fo_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "B_beta_mi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "B_del_mi_c(2,i)"
        I = I + 1:            Data_Strage(I, J) = "B_del_fo_c(2,i)"  '[68]

'******
 '------ 58 [項目名]  B室側 外から３番目の接点
'        i = i + 1:            Data_Strage(i, j) = "Lr_A(0,i)"
'        i = i + 1:            Data_Strage(i, j) = "Lr_A(1,i)"
'        i = i + 1:            Data_Strage(i, j) = "Lr_A(2,i)"
'        i = i + 1:            Data_Strage(i, j) = "Lr_B(2,i)"
'        i = i + 1:            Data_Strage(i, j) = "Lr_B(1,i)"
'        i = i + 1:            Data_Strage(i, j) = "Lr_B(0,i)"
'
'        i = i + 1:            Data_Strage(i, j) = "Lr_A(0,i)"
'        i = i + 1:            Data_Strage(i, j) = "Lr_A(0,i)"
'        i = i + 1:            Data_Strage(i, j) = "Lr_A(0,i)"  '
'******

'------ 69 [項目名]   接点間距離 Ft
        I = I + 1:            Data_Strage(I, J) = "LLt_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "LLt_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "LLt_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Lt_D(j)"          '※
        I = I + 1:            Data_Strage(I, J) = "LLt_B(2, j)"
        I = I + 1:            Data_Strage(I, J) = "LLt_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "LLt_B(0, j)"     '[75]

        I = I + 1:            Data_Strage(I, J) = "Ft_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Ft_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Ft_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Ft_D(j)"
        I = I + 1:            Data_Strage(I, J) = "Ft_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Ft_B(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Ft_AB(j)"       ' [82]'※


'------ 83 [項目名]   接点間距離 Fr
        I = I + 1:            Data_Strage(I, J) = "LLr_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "LLr_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "LLr_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Lr_D(j)"          '※
        I = I + 1:            Data_Strage(I, J) = "LLr_B(2, j)"
        I = I + 1:            Data_Strage(I, J) = "LLr_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "LLr_B(0, j)"

        I = I + 1:            Data_Strage(I, J) = "Fr_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Fr_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Fr_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Fr_D(j)"
        I = I + 1:            Data_Strage(I, J) = "Fr_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Fr_B(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Fr_AB(j)"       '[96]※

'------ 97 [項目名]   接点間距離 and【Moment】  [+15]
        I = I + 1:            Data_Strage(I, J) = "Lmt_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmt_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmt_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmt_D(j)"
        I = I + 1:            Data_Strage(I, J) = "Lmt_B(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmt_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmt_B(0, j)"

        I = I + 1:            Data_Strage(I, J) = "Mmt_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmt_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmt_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmt_D(j)"
        I = I + 1:            Data_Strage(I, J) = "Mmt_B(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmt_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmt_B(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmt_AB(j)"      ' [111]

'------ 112 [項目名]   接点間距離 and【Moment】  [+15]
        I = I + 1:            Data_Strage(I, J) = "Lmr_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmr_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmr_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmr_D(j)"
        I = I + 1:            Data_Strage(I, J) = "Lmr_B(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmr_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Lmr_B(0, j)"

        I = I + 1:            Data_Strage(I, J) = "Mmr_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmr_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmr_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmr_D(j)"
        I = I + 1:            Data_Strage(I, J) = "Mmr_B(2, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmr_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmr_B(0, j)"
        I = I + 1:            Data_Strage(I, J) = "Mmr_AB(j)"      ' [126]


 '------ 127
        I = I + 1:            Data_Strage(I, J) = "Lr_A(0,i)"      '[127]
        I = I + 1:            Data_Strage(I, J) = "Lr_A(1,i)"
        I = I + 1:            Data_Strage(I, J) = "Lr_A(2,i)"
        I = I + 1:            Data_Strage(I, J) = "Lr_B(2,i)"
        I = I + 1:            Data_Strage(I, J) = "Lr_B(1,i)"
        I = I + 1:            Data_Strage(I, J) = "Lr_B(0,i)"      '[132]

        I = I + 1:            Data_Strage(I, J) = "sin A(0,i)"
        I = I + 1:            Data_Strage(I, J) = "sin A(1,i)"
        I = I + 1:            Data_Strage(I, J) = "sin A(2,i)"
        I = I + 1:            Data_Strage(I, J) = "sin B(2,i)"
        I = I + 1:            Data_Strage(I, J) = "sin B(1,i)"
        I = I + 1:            Data_Strage(I, J) = "sin B(0,i)"     '[138]

'------ 139  [項目名]   Center of chamber area
        I = I + 1:            Data_Strage(I, J) = "xg_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "yg_A(0, j)"
        I = I + 1:            Data_Strage(I, J) = "xg_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "yg_A(1, j)"
        I = I + 1:            Data_Strage(I, J) = "xg_A(2, j)"
        I = I + 1:            Data_Strage(I, J) = "yg_A(2, j)"

        I = I + 1:            Data_Strage(I, J) = "xg_B(0, j)"
        I = I + 1:            Data_Strage(I, J) = "yg_B(0, j)"
        I = I + 1:            Data_Strage(I, J) = "xg_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "yg_B(1, j)"
        I = I + 1:            Data_Strage(I, J) = "xg_B(2, j)"
        I = I + 1:            Data_Strage(I, J) = "yg_B(2, j)"     '[150]


'------ 139  [項目名]   Center of chamber area  [+8]
'        i = i + 1:            Data_Strage(i, j) = "xfo xg_f(0, j)"
'        i = i + 1:            Data_Strage(i, j) = "yfo yg_f(0, j)"
'        i = i + 1:            Data_Strage(i, j) = "xfi xg_f(1, j)"
'        i = i + 1:            Data_Strage(i, j) = "yfi yg_f(1, j)"
'
'        i = i + 1:            Data_Strage(i, j) = "xmo xg_m(0, j)"
'        i = i + 1:            Data_Strage(i, j) = "ymo yg_m(0, j)"
'        i = i + 1:            Data_Strage(i, j) = "xmi xg_m(1, j)"
'        i = i + 1:            Data_Strage(i, j) = "ymi yg_m(1, j)"


'  '*** curve_xw(0, i)
'                 tmp_int = 0
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(0, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(0, i)
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(1, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[142]  ' curve_yw(1, i)
'                 tmp_int = 2
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(2, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(2, i)
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(3, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[146]  ' curve_yw(1, i)


'------ 147  [項目名]   Center of chamber area  [+4]
'                 tmp_int = 4
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(4, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(4, i)
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(5, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[150]  ' curve_yw(5, i)
'                 tmp_int = 6
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(6, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(6, i)
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(7, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[154]   ' curve_yw(7, i)
'                 tmp_int = 8
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(8, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(8, i)
'        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(9, i)
'        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[158]   ' curve_yw(9, i)




'--------------------------
'-- Data 部
'--------------------------

    For J = 0 To dw_n    ' the_c(0) to the_c(end= 366)

        I = 1:            Data_Strage(I, J + 3) = J                                '[1] "Index No."       20171031
        I = I + 1:        Data_Strage(I, J + 3) = phi2(J) * 180 / pi               ' "Phi2_[deg]"     20171031
        I = I + 1:        Data_Strage(I, J + 3) = phi2(J)                          ' "Phi2_"          20171031
        I = I + 1:        Data_Strage(I, J + 3) = phi1(J)                          ' "Phi1_"          20171031
        I = I + 1:        Data_Strage(I, J + 3) = the(J)

        I = I + 1:        Data_Strage(I, J + 3) = (the(0) - the(J)) * 180 / pi     ' the(0) - the(j)   20171031
        I = I + 1:        Data_Strage(I, J + 3) = V_a(J) / 1000    '[cc]
        I = I + 1:        Data_Strage(I, J + 3) = V_b(J) / 1000    '[cc]
        I = I + 1:        Data_Strage(I, J + 3) = Sg2_a(0, J) / 1000    '[cc] ' Discharge chamber

        I = I + 1:         Data_Strage(I, J + 3) = Press_A(1, J)
        I = I + 1:         Data_Strage(I, J + 3) = Press_B(1, J)
        I = I + 1:         Data_Strage(I, J + 3) = Press_A(2, J)
        I = I + 1:         Data_Strage(I, J + 3) = Press_B(2, J)
        I = I + 1:         Data_Strage(I, J + 3) = Press_A(0, J)   ' Discharge chamber Pressure


 '------ 15 [Data]  A室側 外から１番目の接点
        I = I + 1:            Data_Strage(I, J + 3) = xfi_c(0, J)     '[15]
        I = I + 1:            Data_Strage(I, J + 3) = yfi_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = xmo_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = ymo_c(0, J)

        I = I + 1:            Data_Strage(I, J + 3) = RR_fi_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = RR_mo_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = beta_fi_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_fi_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_mo_c(0, J)

 '------ 24 [Data]  B室側 外から１番目の接点
        I = I + 1:            Data_Strage(I, J + 3) = xmi_c(0, J)     'B室内壁接点
        I = I + 1:            Data_Strage(I, J + 3) = ymi_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = xfo_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = yfo_c(0, J)

        I = I + 1:            Data_Strage(I, J + 3) = RR_mi_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = RR_fo_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = beta_mi_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_mi_c(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_fo_c(0, J)

 '------ 33 [Data]   A室側 外から２番目の接点
        I = I + 1:            Data_Strage(I, J + 3) = xfi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = yfi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = xmo_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = ymo_c(1, J)

        I = I + 1:            Data_Strage(I, J + 3) = RR_fi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = RR_mo_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = beta_fi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_fi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_mo_c(1, J)

'------ 42 [Data]   B室側 外から２番目の接点
        I = I + 1:            Data_Strage(I, J + 3) = xmi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = ymi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = xfo_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = yfo_c(1, J)

        I = I + 1:            Data_Strage(I, J + 3) = RR_mi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = RR_fo_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = beta_mi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_mi_c(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_fo_c(1, J)

 '------ 51 [Data]   A室側 外から３番目の接点
        I = I + 1:            Data_Strage(I, J + 3) = xfi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = yfi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = xmo_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = ymo_c(2, J)

        I = I + 1:            Data_Strage(I, J + 3) = RR_fi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = RR_mo_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = beta_fi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_fi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_mo_c(2, J)

'------ 60 [Data]   B室側 外から３番目の接点
        I = I + 1:            Data_Strage(I, J + 3) = xmi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = ymi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = xfo_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = yfo_c(2, J)

        I = I + 1:            Data_Strage(I, J + 3) = RR_mi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = RR_fo_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = beta_mi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_mi_c(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = del_fo_c(2, J)    '[68]

'*****
'------ 58 [Data]   B室側 外から３番目の接点
'        i = i + 1:            Data_Strage(i, j + 3) = Lr_A(0, i)
'        i = i + 1:            Data_Strage(i, j + 3) = Lr_A(1, i)
'        i = i + 1:            Data_Strage(i, j + 3) = Lr_A(2, i)
'        i = i + 1:            Data_Strage(i, j + 3) = Lr_B(2, i)
'        i = i + 1:            Data_Strage(i, j + 3) = Lr_B(1, i)
'        i = i + 1:            Data_Strage(i, j + 3) = Lr_B(0, i)
'
'        i = i + 1:            Data_Strage(i, j + 3) = 0
'        i = i + 1:            Data_Strage(i, j + 3) = 0
'        i = i + 1:            Data_Strage(i, j + 3) = 0    '
'*****


      If J > dw_n_PI(2) Then             ' the_c(0) to the_c(2PI) index( 0 to 183)
         GoTo label_Data_Strage_end
      End If

'------ 69  [Data]   接点間距離
        I = I + 1:            Data_Strage(I, J + 3) = LLt_A(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = LLt_A(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = LLt_A(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lt_D(J)         '※
        I = I + 1:            Data_Strage(I, J + 3) = LLt_B(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = LLt_B(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = LLt_B(0, J)

        I = I + 1:            Data_Strage(I, J + 3) = Ft_a(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Ft_a(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Ft_a(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Ft_d(J)
        I = I + 1:            Data_Strage(I, J + 3) = Ft_b(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Ft_b(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Ft_AB(J)        '[82] '※

'------ 83  [Data]   接点間距離 Fr
        I = I + 1:            Data_Strage(I, J + 3) = LLr_A(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = LLr_A(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = LLr_A(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lr_D(J)         '※
        I = I + 1:            Data_Strage(I, J + 3) = LLr_B(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = LLr_B(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = LLr_B(0, J)

        I = I + 1:            Data_Strage(I, J + 3) = Fr_A(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Fr_A(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Fr_A(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Fr_D(J)
        I = I + 1:            Data_Strage(I, J + 3) = Fr_B(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Fr_B(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Fr_AB(J)        '[96] '※

'------  [Data]   接点間距離 and【Moment】  [+14]
        I = I + 1:            Data_Strage(I, J + 3) = Lmt_A(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmt_A(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmt_A(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmt_D(J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmt_B(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmt_B(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmt_B(0, J)

        I = I + 1:            Data_Strage(I, J + 3) = Mmt_A(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmt_A(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmt_A(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmt_D(J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmt_B(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmt_B(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmt_B(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmt_AB(J)       '[111] '※

'------ 112 [Data]   接点間距離 and【Moment】  [+14]
        I = I + 1:            Data_Strage(I, J + 3) = Lmr_A(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmr_A(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmr_A(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmr_D(J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmr_B(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmr_B(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lmr_B(0, J)

        I = I + 1:            Data_Strage(I, J + 3) = Mmr_A(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmr_A(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmr_A(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmr_D(J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmr_B(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmr_B(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmr_B(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = Mmr_AB(J)       '[126] '※

'------ 127  [項目名]

        I = I + 1:            Data_Strage(I, J + 3) = Lr_A(0, J)      '[127] '※
        I = I + 1:            Data_Strage(I, J + 3) = Lr_A(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lr_A(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lr_B(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lr_B(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = Lr_B(0, J)     '

        I = I + 1:            Data_Strage(I, J + 3) = Sin(beta_fi_c(0, J) + del_mo_c(0, J))
        I = I + 1:            Data_Strage(I, J + 3) = Sin(beta_fi_c(1, J) + del_mo_c(1, J))
        I = I + 1:            Data_Strage(I, J + 3) = Sin(beta_fi_c(2, J) + del_mo_c(2, J))
        I = I + 1:            Data_Strage(I, J + 3) = Sin(beta_mi_c(2, J) + del_mi_c(2, J) + pi)
        I = I + 1:            Data_Strage(I, J + 3) = Sin(beta_mi_c(1, J) + del_mi_c(1, J) + pi)
        I = I + 1:            Data_Strage(I, J + 3) = Sin(beta_mi_c(0, J) + del_mi_c(0, J) + pi)     '[138]

'------ 139 [Data]  Center of chamber area】  [+12]
        I = I + 1:            Data_Strage(I, J + 3) = xg_a(0, J)    '[139] (index:外周より中心へ0,1,2)
        I = I + 1:            Data_Strage(I, J + 3) = yg_a(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = xg_a(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = yg_a(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = xg_a(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = yg_a(2, J)

        I = I + 1:            Data_Strage(I, J + 3) = xg_b(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = yg_b(0, J)
        I = I + 1:            Data_Strage(I, J + 3) = xg_b(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = yg_b(1, J)
        I = I + 1:            Data_Strage(I, J + 3) = xg_b(2, J)
        I = I + 1:            Data_Strage(I, J + 3) = yg_b(2, J)     '[150]


'■■Label point
label_Data_Strage_end:

    Next J

'    For J = 0 To div_n      ' div Index 0 to 360


'------ 139 [Data]  Arc line of area that is the center of chamber area 】
'    I = 138
'                tmp_int = 0
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)                 ' xfo(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)                 ' yfo(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)          ' xfi(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)          ' yfi(j)     '[142]
'                tmp_int = 2
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)                 ' xfo(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)                 ' yfo(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)          ' xfi(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)          ' yfi(j)     '[146]
'
'                tmp_int = 4
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)                 ' xfo2(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)                 ' yfo2(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)          ' xfi2(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)          ' yfi2(j)     '[150]
'                tmp_int = 6
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)              '  xmo2(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)              ' ymo2(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)      ' xmi2(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)      ' ymi2(j)     '[154]
'                tmp_int = 8
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)              ' xfo3(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)              ' yfo3(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)      ' xfi3(j)
'        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)      ' yfi3(j)     '[158]
'
'      If J > dw_n_PI(2) Then             ' the_c(0) to the_c(2PI) index( 0 to 183)
'         GoTo label_Data_Strage_end2
'      End If
'
'
'
' '■■Label point
'label_Data_Strage_end2:
'
'    Next J



'-- Data 一括貼付

    Sheets(DataSheetName).Select
'    Sheets(DataSheetName).Cells.ClearContents                    '：Sheet全Cellの数式、文字列をクリア
    Sheets(DataSheetName).Range("C4:FC999").ClearContents         '：指定Cellの数式、文字列をクリア

    I1 = 1                 ' 貼付先の先頭セルの、行と列 (i1, j1)
    J1 = 1
        With Sheets(DataSheetName)
            .Range(Cells(I1, J1), Cells(dw_n + 3 + I1, dw_c + J1)).Value _
                = WorksheetFunction.Transpose(Data_Strage)
        End With

        With Sheets(DataSheetName)
            .Cells(1, 2).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

'            .Cells(1, 3).Value = Format(Now(), "yyyy/MM/DD")       '　"Date
'            .Cells(1, 4).Value = Format(Now(), "HH:mm:ss")         '　"Date
        End With

End Sub



'======================================================== 【M2-      】
'  関数：Calc_Gravity_Center_wrap()
'           *from inner wall of A chamber
'========================================================

Public Sub Calc_Gravity_Center_wrap()

    Dim I As Long, J As Long
    Dim tmp_i As Long
    Dim the_1 As Double:

    Dim Curve_name As String
    Dim tmp_No As Long

    ReDim x_e(dw_n):   ReDim y_e(dw_n)

' for debug of drowing line
    ReDim xfi2(dw_n):   ReDim yfi2(dw_n)
    ReDim xfo2(dw_n):   ReDim yfo2(dw_n)
    ReDim xmi2(dw_n):   ReDim ymi2(dw_n)
    ReDim xmo2(dw_n):   ReDim ymo2(dw_n)

    ReDim xfi3(dw_n):   ReDim yfi3(dw_n)
    ReDim xfo3(dw_n):   ReDim yfo3(dw_n)
    ReDim xmi3(dw_n):   ReDim ymi3(dw_n)
    ReDim xmo3(dw_n):   ReDim ymo3(dw_n)

' A : center of Chamber fiqure
 J = 11
    ReDim xg_f(J, dw_n):          ReDim yg_f(J, dw_n)             ' FS Wrap 図心
      ReDim Sg_f(J, dw_n):                                           ' FS Wrap 面積
    ReDim xg_m(J, dw_n):          ReDim yg_m(J, dw_n)             ' OS Wrap 図心
      ReDim Sg_m(J, dw_n)                                            ' OS Wrap 面積
 J = 10
    ReDim xg2_a(J, dw_n):         ReDim yg2_a(J, dw_n)            ' A Chanmber 図心
      ReDim Sg2_a(J, dw_n):                                           ' A Chanmber 面積
 J = 10
    ReDim xg2_b(J, dw_n):         ReDim yg2_b(J, dw_n)            ' B Chanmber 図心
      ReDim Sg2_b(J, dw_n):                                           ' B Chanmber 面積
'
'      DataSheetName_3 = "DataSheet_D2"

'---------------------------------------------------------
'     N_wrap_a(j) : Number of A Wrap contact points at each index J
'     N_wrap_b(j) : Number of A Wrap contact points at each index J
'     N_wrap_max  : Max Number of Wrap contact points in range

' I = 0
' For I = 0 To 0   '=dw_n   ' 軸回転範囲 4PI = index 366  , 2PI = index 183 , PI = index 90
' For I = 83 To 183   '=dw_n   ' 軸回転範囲 4PI = index 366  , 2PI = index 183 , PI = index 90

          Debug.Print "    ■Rotation index = " & I & Format(Time, "  HH:mm:ss")

'-----------
      For I = 0 To dw_n
            the_1 = the(I) - qq             '
            x_e(I) = Ro * Cos(the_1)
            y_e(I) = Ro * Sin(the_1)
      Next I

    '--time stamp
          Debug.Print " <Calc_Theta_Index_to_Phi_all_0> " & Format(Time, "  HH:mm:ss")

    '--calc Phi_c_fi(J, I) : Pi毎の各接点の伸開角 -
          'ReDim Phi_c_fi(9, dw_n)    ' phi of the contact point from outer to center
            ' DataSheetName = "DataSheet_5"
            ' call Calc_Theta_Index_Number_Read

          DataSheetName = "DataSheet_7"
          Call Calc_Theta_Index_to_Phi_all_0
    Stop

'         Debug.Print (Format(I, "θ000"));
'         If I Mod 10 = 0 Then Debug.Print

'          the_1 = the(358)
'          the_1 = the(358) - pi
'          the_1 = the(359) - pi
''          the_1 = the(358) - pi - 1 / 180 * pi
'          Phi_00 = Fn_Phi2_at_theta(the_1, DataSheetName)
'
'          the_1 = the(358) - 2 * pi
'          Phi_0 = Fn_Phi2_at_theta(the_1, DataSheetName)
'
'          the_1 = the(358) - 2 * pi - 32 / 180 * pi
'          Phi_1 = Fn_Phi2_at_theta(the_1, DataSheetName)
    Stop

    '--Dara Paste -
          DataSheetName = "DataSheet_7"
          Curve_name = "Phi_c_fi_"
          Call Paste_curve_data_Phi_c_fi("I4", Curve_name, Phi_c_fi(), DataSheetName)
     Stop


'-----------
'[tg4 FS]

    tmp_i = dw_n       ' = dw_n ' 軸回転範囲 4PI/3PI/2PI/1PI =  366/273/183/ 90


    Do While (Phi_0 <= FS_in_srt Or Phi_00 <= FS_out_srt)

       Phi_2 = Phi_c_fi(2, I)       ' out end    Phi_0 < Phi_1 < Phi_2
       Phi_00 = Phi_c_fi(3, I)      ' out start
       Phi_1 = Phi_c_fi(3, I)       ' in end
       Phi_0 = Phi_c_fi(4, I)       ' in start

              tmp_i = I
       I = I + 1
    Loop

    ' For I = 0 To tmp_i    ' = dw_n
    For I = 0 To dw_n

       Phi_2 = Phi_c_fi(2, I)       ' out end    Phi_0 < Phi_1 < Phi_2
       Phi_00 = Phi_c_fi(3, I)      ' out start
       Phi_1 = Phi_c_fi(3, I)       ' in end
       Phi_0 = Phi_c_fi(4, I)       ' in start


        If Phi_1 <= FS_in_srt Or Phi_2 <= FS_out_srt Then
              xg_f(4, I) = 0
              yg_f(4, I) = 0
              Sg_f(4, I) = 0

        ElseIf Phi_0 <= FS_in_srt Or Phi_00 <= FS_out_srt Then
              Phi_00 = Phi_c_fi(3, tmp_i)
              Phi_0 = Phi_c_fi(4, tmp_i)

          Call Get_Gravity_Center_FS_wrap(Phi_0, Phi_1, Phi_00, Phi_2)
              xg_f(4, I) = xg_a_tmp
              yg_f(4, I) = yg_a_tmp
              Sg_f(4, I) = Area_tmp

        Else

          Call Get_Gravity_Center_FS_wrap(Phi_0, Phi_1, Phi_00, Phi_2)
              xg_f(4, I) = xg_a_tmp
              yg_f(4, I) = yg_a_tmp
              Sg_f(4, I) = Area_tmp

        End If

          ' Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

        '--Dara Paste -
           Curve_name = "tg4 FS_":    ' DataSheetName = "DataSheet_6"
'           Call Paste_curve_data_Num(0, Curve_name & "in", x_in(), y_in(), DataSheetName)
'           Call Paste_curve_data_Num(1, Curve_name & "out", x_out(), y_out(), DataSheetName)
'           Call Paste_curve_data_Num(2, Curve_name & "in", x_in(), y_in(), DataSheetName)
'           Call Paste_curve_data_Num(3, Curve_name & "out", x_out(), y_out(), DataSheetName)
'           Stop

    Next I

  '--Dara Paste -
        Curve_name = "tg4_":
        tmp_No = 4:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tg_No("U4", Curve_name, tmp_No, DataSheetName)
     Stop


'[tg3 FS]
    For I = 0 To tmp_i     ' = dw_n

       Phi_2 = Phi_c_fi(3, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(4, I)
       Phi_0 = Phi_c_fi(5, I)

          Call Get_Gravity_Center_FS_wrap(Phi_0, Phi_1, Phi_1, Phi_2)
              xg_f(3, I) = xg_a_tmp
              yg_f(3, I) = yg_a_tmp
              Sg_f(3, I) = Area_tmp
          ' Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
      ' Curve_name = "tg3 FS_":    ' DataSheetName = "DataSheet_2"
      ' Call Paste_curve_data_Num(2, Curve_name & "in", x_in(), y_in(), DataSheetName)
      ' Call Paste_curve_data_Num(3, Curve_name & "out", x_out(), y_out(), DataSheetName)
      ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tg3_":
        tmp_No = 3:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tg_No("X4", Curve_name, tmp_No, DataSheetName)
     Stop



'[tg2 FS]
    For I = 0 To tmp_i     ' = dw_n

       Phi_2 = Phi_c_fi(4, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(5, I)
       Phi_0 = Phi_c_fi(6, I)

      If (Phi_2 >= 0 And Phi_1 >= 0 And Phi_0 >= 0) Then
          Call Get_Gravity_Center_FS_wrap(Phi_0, Phi_1, Phi_1, Phi_2)
              xg_f(2, I) = xg_a_tmp
              yg_f(2, I) = yg_a_tmp
              Sg_f(2, I) = Area_tmp
          ' Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(1, x_in, y_in)
      Else
              xg_f(2, I) = 0
              yg_f(2, I) = 0
              Sg_f(2, I) = 0
      End If

      '--Dara Paste -
      ' Curve_name = "tg2 FS_":    ' DataSheetName = "DataSheet_2"
      ' Call Paste_curve_data_Num(4, Curve_name & "in", x_in(), y_in(), DataSheetName)
      ' Call Paste_curve_data_Num(5, Curve_name & "out", x_out(), y_out(), DataSheetName)
      ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tg2_":
        tmp_No = 2:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tg_No("AA4", Curve_name, tmp_No, DataSheetName)
     Stop



'[tg1 FS]
    For I = 0 To tmp_i     ' = dw_n

       Phi_2 = Phi_c_fi(5, I)       ' outer wall  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(6, I)       ' inner wall  Phi_0 < Phi_1 < Phi_2      PI

       Phi_00 = Wrap_Start_angle_min(2)   ' outer = 106-6 wall  Phi_0 < Phi_1 < Phi_2      PI
       Phi_0 = Wrap_Start_angle_min(1)   ' inner  =106-56

        '  Phi_00 = ((110) * pi / 180)   ' outer = 106-6 wall  Phi_0 < Phi_1 < Phi_2      PI
        '  Phi_0 = ((60) * pi / 180)     ' inner  =106-56

   If (Phi_2 > Phi_00 And Phi_1 > Phi_0) Then
          Call Get_Gravity_Center_FS_wrap(Phi_0, Phi_1, Phi_00, Phi_2)
                xg_f(1, I) = xg_a_tmp
                yg_f(1, I) = yg_a_tmp
                Sg_f(1, I) = Area_tmp

    ElseIf (Phi_c_fi(4, I) > Phi_00 And Phi_c_fi(5, I) > Phi_0) Then
            Phi_2 = Phi_c_fi(4, I)       ' outer wall  Phi_0 < Phi_1 < Phi_2      PI
            Phi_1 = Phi_c_fi(5, I)       ' inner wall  Phi_0 < Phi_1 < Phi_2      PI

          Call Get_Gravity_Center_FS_wrap(Phi_0, Phi_1, Phi_00, Phi_2)
                xg_f(1, I) = xg_a_tmp
                yg_f(1, I) = yg_a_tmp
                Sg_f(1, I) = Area_tmp
    Else
'          Stop
                xg_f(1, I) = 0
                yg_f(1, I) = 0
                Sg_f(1, I) = 0
    End If

          ' Call change_Wrap_data_to_curve_xw(0, x_out, y_out)      '(0, xfo + dx, yfo + dy)
          ' Call change_Wrap_data_to_curve_xw(1, x_in, y_in)        '(0, xfi + dx, yfi + dy)

          '--Dara Paste -
          ' Curve_name = "tg1 FS_":    ' DataSheetName = "DataSheet_2"
          ' Call Paste_curve_data_Num(6, Curve_name & "in", x_in(), y_in(), DataSheetName)
          ' Call Paste_curve_data_Num(7, Curve_name & "out", x_out(), y_out(), DataSheetName)
          ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tg1_":
        tmp_No = 1:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tg_No("AD4", Curve_name, tmp_No, DataSheetName)
     Stop

  '   GoTo Label_Gravity_center_end


'----------
'[tp6 OS]
    For I = 0 To tmp_i     ' = dw_n

        If I >= 90 Then
            Phi_2 = Phi_c_fi(1, 0)       '  Phi_0 < Phi_1 < Phi_2      PI
            Phi_00 = Phi_c_fi(2, I)
            Phi_1 = Phi_c_fi(2, 0)
            Phi_0 = Phi_c_fi(3, I)
        ElseIf (90 > I And I > 0) Then
            Phi_2 = Phi_c_fi(1, 0)       '  Phi_0 < Phi_1 < Phi_2      PI
            Phi_00 = Phi_c_fi(1, I)
            Phi_1 = Phi_c_fi(2, 0)
            Phi_0 = Phi_c_fi(2, I)
        ElseIf I = 0 Then
            Phi_2 = 0      ' Phi_c_fi(1, 0)       '  Phi_0 < Phi_1 < Phi_2      PI
            Phi_00 = 0     ' Phi_c_fi(2, i)
            Phi_1 = 0      ' Phi_c_fi(2, 0)
            Phi_0 = 0      ' Phi_c_fi(3, i)
       Else
           Stop
       End If

    ' <M2-1-F1>
        If (Phi_2 = Phi_00) And (Phi_1 = Phi_0) Then
            Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
                xg_m(6, I) = 0   ' + Ro * Cos(the_1)
                yg_m(6, I) = 0  '  + Ro * Sin(the_1)
                Sg_m(6, I) = 0
        Else
            Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
                xg_m(6, I) = xg_a_tmp  '  + Ro * Cos(the_1)
                yg_m(6, I) = yg_a_tmp  '  + Ro * Sin(the_1)
                Sg_m(6, I) = Area_tmp
        End If

          ' Call change_Wrap_data_to_curve_xw(2, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(3, x_in, y_in)

      '--Dara Paste -
          ' Curve_name = "tp6 OS_":    ' DataSheetName = "DataSheet_2"
          ' Call Paste_curve_data_Num(8, Curve_name & "in", x_in(), y_in(), DataSheetName)
          ' Call Paste_curve_data_Num(9, Curve_name & "out", x_out(), y_out(), DataSheetName)
          ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tp6_":
        tmp_No = 6:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tp_No("AG4", Curve_name, tmp_No, DataSheetName)
     Stop


'GoTo Label_Gravity_center_end

'[tp5 OS]
    For I = 0 To tmp_i     ' = dw_n

       Phi_2 = Phi_c_fi(1, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(2, I)
       Phi_0 = Phi_c_fi(3, I)

          Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_1, Phi_2, the_1)
                xg_m(5, I) = xg_a_tmp  ' + Ro * Cos(the_1)
                yg_m(5, I) = yg_a_tmp  ' + Ro * Sin(the_1)
                Sg_m(5, I) = Area_tmp
          ' Call change_Wrap_data_to_curve_xw(2, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(3, x_in, y_in)

      '--Dara Paste -
          ' Curve_name = "tp5 OS_":    ' DataSheetName = "DataSheet_2"
          ' Call Paste_curve_data_Num(8, Curve_name & "in", x_in(), y_in(), DataSheetName)
          ' Call Paste_curve_data_Num(9, Curve_name & "out", x_out(), y_out(), DataSheetName)
          ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tp5_":
        tmp_No = 5:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tp_No("AJ4", Curve_name, tmp_No, DataSheetName)
     Stop


'[tp4 OS]
    For I = 0 To tmp_i     ' = dw_n

       Phi_2 = Phi_c_fi(2, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(3, I)
       Phi_0 = Phi_c_fi(4, I)

          Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_1, Phi_2, the_1)
              xg_m(4, I) = xg_a_tmp   ' + Ro * Cos(the_1)
              yg_m(4, I) = yg_a_tmp   ' + Ro * Sin(the_1)
              Sg_m(4, I) = Area_tmp
          ' Call change_Wrap_data_to_curve_xw(2, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(3, x_in, y_in)

      '--Dara Paste -
          ' Curve_name = "tp4 OS_":    ' DataSheetName = "DataSheet_2"
          ' Call Paste_curve_data_Num(10, Curve_name & "in", x_in(), y_in(), DataSheetName)
          ' Call Paste_curve_data_Num(11, Curve_name & "out", x_out(), y_out(), DataSheetName)
          ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tp4_":
        tmp_No = 4:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tp_No("AM4", Curve_name, tmp_No, DataSheetName)
     Stop


'[tp3 OS]
    For I = 0 To tmp_i     ' = dw_n

       Phi_2 = Phi_c_fi(3, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(4, I)
       Phi_0 = Phi_c_fi(5, I)

          Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_1, Phi_2, the_1)
              xg_m(3, I) = xg_a_tmp  '  + Ro * Cos(the_1)
              yg_m(3, I) = yg_a_tmp  '  + Ro * Sin(the_1)
              Sg_m(3, I) = Area_tmp
          ' Call change_Wrap_data_to_curve_xw(2, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(3, x_in, y_in)

      '--Dara Paste -
          ' Curve_name = "tp3 OS_":    ' DataSheetName = "DataSheet_2"
          ' Call Paste_curve_data_Num(12, Curve_name & "in", x_in(), y_in(), DataSheetName)
          ' Call Paste_curve_data_Num(13, Curve_name & "out", x_out(), y_out(), DataSheetName)
          ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tp3_":
        tmp_No = 3:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tp_No("AP4", Curve_name, tmp_No, DataSheetName)
     Stop


'[tp2 OS]
    For I = 0 To tmp_i     ' = dw_n

       Phi_2 = Phi_c_fi(4, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(5, I)
       Phi_0 = Phi_c_fi(6, I)

          Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_1, Phi_2, the_1)
              xg_m(2, I) = xg_a_tmp  '  + Ro * Cos(the_1)
              yg_m(2, I) = yg_a_tmp  '  + Ro * Sin(the_1)
              Sg_m(2, I) = Area_tmp
          ' Call change_Wrap_data_to_curve_xw(2, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(3, x_in, y_in)

      '--Dara Paste -
          ' Curve_name = "tp2 OS_":    ' DataSheetName = "DataSheet_2"
          ' Call Paste_curve_data_Num(14, Curve_name & "in", x_in(), y_in(), DataSheetName)
          ' Call Paste_curve_data_Num(15, Curve_name & "out", x_out(), y_out(), DataSheetName)
          ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tp2_":
        tmp_No = 2:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tp_No("AS4", Curve_name, tmp_No, DataSheetName)
     Stop



'[tp1 OS]
    For I = 0 To tmp_i     ' = dw_n

       Phi_2 = Phi_c_fi(5, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(6, I)

       Phi_0 = Wrap_Start_angle_min(3)   ' inner  =106-86
       Phi_00 = Wrap_Start_angle_min(4)  ' outer  =142-12  wall  Phi_0 < Phi_1 < Phi_2      PI

       OS_in_srt_0 = Phi_0          ' add 20180711
       OS_out_srt_0 = Phi_00

    If Phi_0 >= Phi_1 Then
       Phi_2 = Phi_c_fi(4, I)     '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(5, I)
    End If

    If (Phi_2 > 0 And Phi_1 > 0 And Phi_0 > 0 And Phi_00 > 0) Then

          Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
            xg_m(1, I) = xg_a_tmp  '  + Ro * Cos(the_1)
            yg_m(1, I) = yg_a_tmp  '  + Ro * Sin(the_1)
            Sg_m(1, I) = Area_tmp
          ' Call change_Wrap_data_to_curve_xw(2, x_out, y_out)
          ' Call change_Wrap_data_to_curve_xw(3, x_in, y_in)
    Else
'          Stop
            xg_m(1, I) = 0
            yg_m(1, I) = 0
            Sg_m(1, I) = 0

    End If

      '--Dara Paste -
          ' Curve_name = "tp1 OS_":    ' DataSheetName = "DataSheet_2"
          ' Call Paste_curve_data_Num(16, Curve_name & "in", x_in(), y_in(), DataSheetName)
          ' Call Paste_curve_data_Num(17, Curve_name & "out", x_out(), y_out(), DataSheetName)
          ' Stop

    Next I

  '--Dara Paste -
        Curve_name = "tp1_":
        tmp_No = 1:             ' DataSheetName = "DataSheet_7"
        Call Paste_curve_data_tp_No("AV4", Curve_name, tmp_No, DataSheetName)
     Stop


     Debug.Print "    tp1 OS " & Format(Time, "  HH:mm:ss")



'Label_Data_Strage_Paste_D_chamber:

'-----------
'    D Chamber : Discharge --------
'-----------
'[pp1-Pg(1)]  ' halfmoon_A + halfmoon_B -tp1 -tg1

      ' For tmp_Num2 = 0 To dw_n_PI(2)           'div_n     ' dw_n_PI(2)=183
      '     i =  tmp_Num2
      ' Next  tmp_Num2
      '
      '    DataSheetName = "DataSheet_D2"
      '    Call Data_Strage_to_array_4
      'Stop

'       Phi_2 = Phi_c_fi(4, I)       '  Phi_0 < Phi_1 < Phi_2      PI
       Phi_1 = Phi_c_fi(5, I)
       Phi_0 = Phi_c_fi(6, I)

    If Phi_0 > 0 Then

     '<A side FS_in pp1>   half moon_A
          Call Get_Gravity_Center_Chamber_D_FS(Phi_0, Phi_1, the_1)
                xg_f(0, I) = xg_a_tmp
                yg_f(0, I) = yg_a_tmp
                Sg_f(0, I) = Area_tmp

          Call change_Wrap_data_to_curve_xw(4, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(5, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pp1 FS_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(0, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(1, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

    ' <B side OS_in tp0>  half moon_B
          Call Get_Gravity_Center_Chamber_D_OS(Phi_0, Phi_1, the_1)
                xg_m(0, I) = xg_a_tmp   '+ Ro * Cos(the_1)
                yg_m(0, I) = yg_a_tmp   '+ Ro * Sin(the_1)
                Sg_m(0, I) = Area_tmp

          Call change_Wrap_data_to_curve_xw(6, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(7, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tp1 OS_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(2, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(3, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

  '-----------
  '  Discharge chamber
  '    [pp1-Pg(1)]  ' halfmoon_A + halfmoon_B -tp1 -tg1
  '-----------
          Sg2_a(0, I) = Sg_f(0, I) + Sg_m(0, I) - Sg_f(1, I) - Sg_m(1, I)

          xg2_a(0, I) = (Sg_f(0, I) * xg_f(0, I) + Sg_m(0, I) * xg_m(0, I) _
                        - Sg_f(1, I) * xg_f(1, I) - Sg_m(1, I) * xg_m(1, I)) / Sg2_a(0, I)

          yg2_a(0, I) = (Sg_f(0, I) * yg_f(0, I) + Sg_m(0, I) * yg_m(0, I) _
                        - Sg_f(1, I) * yg_f(1, I) - Sg_m(1, I) * yg_m(1, I)) / Sg2_a(0, I)
   Else
          xg_f(0, I) = 0
          yg_f(0, I) = 0
          Sg_f(0, I) = 0
          xg_m(0, I) = 0
          yg_m(0, I) = 0
          Sg_m(0, I) = 0

          Sg2_a(0, I) = 0
          xg2_a(0, I) = 0
          yg2_a(0, I) = 0
   End If


'   GoTo Label_Gravity_center_end

'-----------
'--- A Chamber --------
'-----------
'[Pp(3) : A Chamber Fs_in-OS_out]

       Phi_2 = Phi_c_fi(1, I)       '   Phi_1 < Phi_2      2PI
       Phi_1 = Phi_c_fi(3, I)

          Call Get_Gravity_Center_Chamber_A(Phi_1, Phi_2)
              xg2_a(3, I) = xg_a_tmp
              yg2_a(3, I) = yg_a_tmp
              Sg2_a(3, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pp3 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(4, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(5, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

'[Pp(2) : A Chamber Fs_in-OS_out]
       Phi_2 = Phi_c_fi(3, I)       '   Phi_1 < Phi_2      2PI
       Phi_1 = Phi_c_fi(5, I)

          Call Get_Gravity_Center_Chamber_A(Phi_1, Phi_2)
              xg2_a(2, I) = xg_a_tmp
              yg2_a(2, I) = yg_a_tmp
              Sg2_a(2, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(2, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(3, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pp2 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(6, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(7, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

'--- B Chamber --------

'[Pg(2) : B Chamber OS_in-FS_ou]
       Phi_2 = Phi_c_fi(3, I)       '   Phi_1 < Phi_2      2PI
       Phi_1 = Phi_c_fi(5, I)

          Call Get_Gravity_Center_Chamber_B(Phi_1, Phi_2)
              xg2_b(2, I) = xg_a_tmp
              yg2_b(2, I) = yg_a_tmp
              Sg2_b(2, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(4, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(5, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg2 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(8, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(9, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

'-----------
' [Pg(3) : Suction Chamber ] --- Pre Suction Chamber : suction inlet
'
'          S = area1 + area2 + area3 - OSwrap
'              clecent-start / clecent-inlet / suction-inlet / tp6
'
'-----------
        '  Phi_0 = Phi_c_fi(3, i)        ' Phi_0 < Phi_1   No.2-3_contact_Point
        '  Phi_1 = Phi_c_fi(2, i)
        '  Phi_00 = Phi_c_fi(3, i)       ' Phi_00 < Phi_2  No.2-3_contact_Point
        '  Phi_2 = Phi_c_fi(2, i)

  'i = 138: the_1 = the(i) - qq

   If (the_c(I) = 0 Or the_c(I) = 2 * pi) Then

     '-----------
     ' area1
        Phi_00 = Phi_c_fi(3, I)       ' OS_in   : Phi_00 < Phi_2  No.2-3_contact_Point
           Phi_2 = Phi_c_fi(2, 0)
        Phi_0 = Phi_c_fi(3, I)        ' FS_out  : Phi_0 < Phi_1   No.2-3_contact_Point
           Phi_1 = Phi_c_fi(2, 0)

        Call Get_Gravity_Center_Chamber_Suction_1B(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg2_b(4, I) = xg_a_tmp
              yg2_b(4, I) = yg_a_tmp
              Sg2_b(4, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(6, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(7, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area1 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(0, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(1, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

    'GoTo Label_Gravity_center_end

     '-----------
     ' area2
          xg2_b(5, I) = 0
          yg2_b(5, I) = 0
          Sg2_b(5, I) = 0
     '-----------
     ' area2
        Phi_00 = Phi_c_fi(1, 0)       ' FS-in_end  Phi_2 ～ inlet point
           Phi_2 = angle_Rfi_c
        Phi_0 = Phi_c_fi(2, 0)        ' FS-out_end  Phi_1 ～ inlet point
           Phi_1 = angle_Rfo_c

      'Suction inlet
        Call Get_Gravity_Center_Chamber_Suction_3(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg2_b(6, I) = xg_a_tmp
              yg2_b(6, I) = yg_a_tmp
              Sg2_b(6, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(8, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(9, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 inlet":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(2, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(3, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop
'
     '-----------
     ' Gravity Center area1-area3
        Sg2_b(3, I) = Sg2_b(4, I) + Sg2_b(5, I) + Sg2_b(6, I)

        xg2_b(3, I) = (Sg2_b(4, I) * xg2_b(4, I) + Sg2_b(5, I) * xg2_b(5, I) _
                                          + Sg2_b(6, I) * xg2_b(6, I)) / Sg2_b(3, I)
        yg2_b(3, I) = (Sg2_b(4, I) * yg2_b(4, I) + Sg2_b(5, I) * yg2_b(5, I) _
                                          + Sg2_b(6, I) * yg2_b(6, I)) / Sg2_b(3, I)

   ' [Pg(3)-case2
   ElseIf (0 < the_c(I) And the_c(I) < pi) Then        ' 0 < the_c(i) < 180

     '-----------
     ' area1   B/2-chamber
        Phi_00 = Phi_c_fi(3, I)       ' FS_out   : Phi_00 < Phi_2  No.2-3_contact_Point
           Phi_2 = Phi_c_fi(2, I)
        Phi_0 = Phi_c_fi(3, I)        ' FS_in  : Phi_0 < Phi_1   No.2-3_contact_Point
           Phi_1 = Phi_c_fi(2, I)

        Call Get_Gravity_Center_Chamber_Suction_1B(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg2_b(4, I) = xg_a_tmp
              yg2_b(4, I) = yg_a_tmp
              Sg2_b(4, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(6, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(7, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area1 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(0, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(1, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

    'GoTo Label_Gravity_center_end

     '-----------
     ' area2   AB-area (Between FS-in and FS-out)
        Phi_00 = Phi_c_fi(1, I)       ' FS_out   : Phi_00 < Phi_2  No.2-3_contact_Point
           Phi_2 = Phi_c_fi(1, 0)
        Phi_0 = Phi_c_fi(2, I)        ' FS_in  : Phi_0 < Phi_1   No.2-3_contact_Point
           Phi_1 = Phi_c_fi(2, 0)

        Call Get_Gravity_Center_Chamber_Suction_2(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg2_b(5, I) = xg_a_tmp
              yg2_b(5, I) = yg_a_tmp
              Sg2_b(5, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(8, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(9, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area2 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(2, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(3, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

    'GoTo Label_Gravity_center_end

     '-----------
     ' area3
        Phi_00 = Phi_c_fi(1, 0)       ' FS-in_end  Phi_2 ～ inlet point
           Phi_2 = angle_Rfi_c
        Phi_0 = Phi_c_fi(2, 0)        ' FS-out_end  Phi_1 ～ inlet point
           Phi_1 = angle_Rfo_c

        Call Get_Gravity_Center_Chamber_Suction_3(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
                xg2_b(6, I) = xg_a_tmp
                yg2_b(6, I) = yg_a_tmp
                Sg2_b(6, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area3 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(4, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(5, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

    'GoTo Label_Gravity_center_end

     '-----------
     ' OS wrap end-part
        Phi_00 = Phi_c_fi(1, I)       ' FS-in_end  Phi_2 ～ inlet point
           Phi_2 = Phi_c_fi(1, 0)
        Phi_0 = Phi_c_fi(2, I)        ' FS-out_end  Phi_1 ～ inlet point
           Phi_1 = Phi_c_fi(2, 0)

        Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg_m(6, I) = xg_a_tmp
              yg_m(6, I) = yg_a_tmp
              Sg_m(6, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 end OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(6, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(7, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

     '-----------
         Sg2_b(3, I) = Sg2_b(4, I) + Sg2_b(5, I) + Sg2_b(6, I) - Sg_m(6, I)

         xg2_b(3, I) = (Sg2_b(4, I) * xg2_b(4, I) + Sg2_b(5, I) * xg2_b(5, I) _
                        + Sg2_b(6, I) * xg2_b(6, I) - Sg_m(6, I) * xg_m(6, I)) / Sg2_b(3, I)

         yg2_b(3, I) = (Sg2_b(4, I) * yg2_b(4, I) + Sg2_b(5, I) * yg2_b(5, I) _
                        + Sg2_b(6, I) * yg2_b(6, I) - Sg_m(6, I) * yg_m(6, I)) / Sg2_b(3, I)

    'GoTo Label_Gravity_center_end

   ' [Pg(3)-case3
   ElseIf (the_c(I) = pi) Then                '  the_c(i) = 180

     '-----------
     ' area1   B/2-chamber
        Phi_00 = Phi_c_fi(4, 0)       ' FS_out   : Phi_00 < Phi_2  No.2-3_contact_Point
           Phi_2 = Phi_c_fi(3, 0)
        Phi_0 = Phi_c_fi(4, 0)        ' FS_in  : Phi_0 < Phi_1   No.2-3_contact_Point
           Phi_1 = Phi_c_fi(3, 0)

        Call Get_Gravity_Center_Chamber_Suction_1B(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg2_b(4, I) = xg_a_tmp
              yg2_b(4, I) = yg_a_tmp
              Sg2_b(4, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area1 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(0, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(1, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

     '-----------
     ' area2   AB-area (Between FS-in and FS-out)
        Phi_00 = Phi_c_fi(2, 0)       ' FS_out   : Phi_00 < Phi_2  No.2-3_contact_Point
           Phi_2 = Phi_c_fi(1, 0)
        Phi_0 = Phi_c_fi(3, 0)        ' FS_in  : Phi_0 < Phi_1   No.2-3_contact_Point
           Phi_1 = Phi_c_fi(2, 0)

        Call Get_Gravity_Center_Chamber_Suction_2(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg2_b(5, I) = xg_a_tmp
              yg2_b(5, I) = yg_a_tmp
              Sg2_b(5, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area2 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(2, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(3, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

     '-----------
     ' area3
        Phi_00 = Phi_c_fi(1, 0)       ' FS-in_end  Phi_2 ～ inlet point
           Phi_2 = angle_Rfi_c
        Phi_0 = Phi_c_fi(2, 0)        ' FS-out_end  Phi_1 ～ inlet point
           Phi_1 = angle_Rfo_c

        Call Get_Gravity_Center_Chamber_Suction_3(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg2_b(6, I) = xg_a_tmp
              yg2_b(6, I) = yg_a_tmp
              Sg2_b(6, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area3 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(4, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(5, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

     '-----------
     ' OS wrap end-part
        Phi_00 = Phi_c_fi(2, 0)       ' FS-in_end  Phi_2 ～ inlet point
           Phi_2 = Phi_c_fi(1, 0)
        Phi_0 = Phi_c_fi(3, 0)        ' FS-out_end  Phi_1 ～ inlet point
           Phi_1 = Phi_c_fi(2, 0)

        Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg_m(6, I) = xg_a_tmp
              yg_m(6, I) = yg_a_tmp
              Sg_m(6, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 end OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(6, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(7, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop


    'GoTo Label_Gravity_center_end

         Sg2_b(3, I) = Sg2_b(4, I) + Sg2_b(5, I) + Sg2_b(6, I) - Sg_m(6, I)

         xg2_b(3, I) = (Sg2_b(4, I) * xg2_b(4, I) + Sg2_b(5, I) * xg2_b(5, I) _
                        + Sg2_b(6, I) * xg2_b(6, I) - Sg_m(6, I) * xg_m(6, I)) / Sg2_b(3, I)

         yg2_b(3, I) = (Sg2_b(4, I) * yg2_b(4, I) + Sg2_b(5, I) * yg2_b(5, I) _
                        + Sg2_b(6, I) * yg2_b(6, I) - Sg_m(6, I) * yg_m(6, I)) / Sg2_b(3, I)

   ' [Pg(3)-case4
   ElseIf (pi < the_c(I) And the_c(I) < 2 * pi) Then       ' 180 < the_c(i) < 360

     '-----------
     ' area1   A/2-chamber
        Phi_00 = Phi_c_fi(1, I)       ' FS_in   : Phi_00 < Phi_2  No.2-3_contact_Point
           Phi_2 = Phi_c_fi(9, I)
        Phi_0 = Phi_c_fi(1, I)        ' OS_out  : Phi_0 < Phi_1   No.2-3_contact_Point
           Phi_1 = Phi_c_fi(9, I)

        Call Get_Gravity_Center_Chamber_Suction_1A(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
                xg2_b(4, I) = xg_a_tmp
                yg2_b(4, I) = yg_a_tmp
                Sg2_b(4, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

       '--Dara Paste -
        Curve_name = "pg3 area1 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(0, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(1, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop


    'GoTo Label_Gravity_center_end

     '-----------
     ' area2   AB-area (Between FS-in and FS-out)
        Phi_00 = Phi_c_fi(9, I)       ' FS_in   : Phi_00 < Phi_2  No.2-3_contact_Point
           Phi_2 = Phi_c_fi(1, 0)
        Phi_0 = Phi_c_fi(8, I)        ' FS_out  : Phi_0 < Phi_1   No.2-3_contact_Point
           Phi_1 = Phi_c_fi(2, 0)

        Call Get_Gravity_Center_Chamber_Suction_2(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
                xg2_b(5, I) = xg_a_tmp
                yg2_b(5, I) = yg_a_tmp
                Sg2_b(5, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area2 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(2, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(3, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop


    'GoTo Label_Gravity_center_end

     '-----------
     ' area3
        Phi_00 = Phi_c_fi(1, 0)       ' FS-in_end  Phi_2 ～ inlet point
           Phi_2 = angle_Rfi_c
        Phi_0 = Phi_c_fi(2, 0)        ' FS-out_end  Phi_1 ～ inlet point
           Phi_1 = angle_Rfo_c

        Call Get_Gravity_Center_Chamber_Suction_3(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg2_b(6, I) = xg_a_tmp
              yg2_b(6, I) = yg_a_tmp
              Sg2_b(6, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 area3 OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(4, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(5, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

     '-----------
     ' OS wrap end-part
        Phi_00 = Phi_c_fi(9, I)       ' FS-in_end  Phi_2 ～ inlet point
           Phi_2 = Phi_c_fi(1, 0)
        Phi_0 = Phi_c_fi(8, I)        ' FS-out_end  Phi_1 ～ inlet point
           Phi_1 = Phi_c_fi(2, 0)

        Call Get_Gravity_Center_OS_wrap(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
              xg_m(6, I) = xg_a_tmp
              yg_m(6, I) = yg_a_tmp
              Sg_m(6, I) = Area_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "pg3 emd OS":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(6, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(7, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

     '---------------------------------------

         Sg2_b(3, I) = Sg2_b(4, I) + Sg2_b(5, I) + Sg2_b(6, I) - Sg_m(6, I)

         xg2_b(3, I) = (Sg2_b(4, I) * xg2_b(4, I) + Sg2_b(5, I) * xg2_b(5, I) _
                        + Sg2_b(6, I) * xg2_b(6, I) - Sg_m(6, I) * xg_m(6, I)) / Sg2_b(3, I)

         yg2_b(3, I) = (Sg2_b(4, I) * yg2_b(4, I) + Sg2_b(5, I) * yg2_b(5, I) _
                        + Sg2_b(6, I) * yg2_b(6, I) - Sg_m(6, I) * yg_m(6, I)) / Sg2_b(3, I)

   ' [Pg(3)-case5
   Else
      Stop
   End If

'Stop

  'GoTo Label_Gravity_center_end


'-----------
'[PW : oil groove ]    a:outer/ b:inner
'-----------
        '  arc1 : From start side to conection side
        '  arc2 : From connection side to inlet side
        '-- right outer arc1 spec.   '-- left outer arc2 spec.
        '   r1_oilgroove = 53.72      r2_oilgroove = 56.35
        '   x1_oilgroove_c = 0.89     x2_oilgroove_c = 1.05
        '   y1_oilgroove_c = 2.4      y2_oilgroove_c = -0.25
        '   t1_oilgroove = 1.5        t2_oilgroove = t1_oilgroove
        '   angle1_oilgroove_0             angle2_oilgroove_2
        '    = start angel (-70)*pi/180     = end angle (180 + 61)*pi/180

  '[PW-1]  arc1,arc2の中心の位置関係
  '-- Length of between arc1-center and arc2-center
      L_tmp0 = Sqr((x1_oilgroove_c - x2_oilgroove_c) ^ 2 + (y1_oilgroove_c - y2_oilgroove_c) ^ 2)

  '-- angle of L_tmp0 on the FS xy-cordinate
      angle_tmp0 = Atan2((y2_oilgroove_c - y1_oilgroove_c), (x2_oilgroove_c - x1_oilgroove_c))

    ' chage the range : 0 < Atan2 < 2π ==> -π < Atan2 < π
        If pi <= angle_tmp0 Then
          angle_tmp0 = angle_tmp0 - 2 * pi
        End If

  '[PW-2]　接続点の扇形開口角度
  '--- on oilgroove arc1 range
    ' connection point of angle : arc1 and arc2  [wtih hte the law of cosines 余弦定理]

    ' a) outer arc1a  b) inner arc1b
      angle_tmp1a = arcCos(((L_tmp0) ^ 2 + (r1_oilgroove) ^ 2 - (r2_oilgroove) ^ 2) _
                    / (2 * L_tmp0 * r1_oilgroove))

      angle_tmp1b = arcCos(((L_tmp0) ^ 2 + (r1_oilgroove - t1_oilgroove) ^ 2 _
          - (r2_oilgroove - t2_oilgroove) ^ 2) / (2 * L_tmp0 * (r1_oilgroove - t1_oilgroove)))

  '--- on oilgroove arc2 range
    '    a:outer arc2a  b:inner arc2b
      angle_tmp2a = arcCos(((L_tmp0) ^ 2 + (r2_oilgroove) ^ 2 - (r1_oilgroove) ^ 2) _
                    / (2 * L_tmp0 * r2_oilgroove))

      angle_tmp2b = arcCos(((L_tmp0) ^ 2 + (r2_oilgroove - t2_oilgroove) ^ 2 _
         - (r1_oilgroove - t1_oilgroove) ^ 2) / (2 * L_tmp0 * (r2_oilgroove - t2_oilgroove)))

      '      ' oilgroove arc1 of start angle
      '           angle1_oilgroove_0 = (-70) * pi / 180          ' start angle      '
      '      ' oilgroove arc2 of end angle
      '           angle2_oilgroove_2 = (180 + 61) * pi / 180     ' end angle

    ' arc1,arc2に交点が無い場合の処理
        If (Abs(r1_oilgroove + r2_oilgroove) < L_tmp0) _
                          And (L_tmp0 < Abs(r1_oilgroove - r2_oilgroove)) Then
              Stop
        End If

  '[PW-3]--- arc Connection angle  a:outer arc1a  b:inner arc1b

   If ((0 <= angle_tmp0) And (angle_tmp0 < pi / 2)) Then            ' 0 =< angle <PI/2
      '--- oilgroove arc1 range  out 0-1 and inner 0-2
         angle1_oilgroove_1 = (angle_tmp0 + angle_tmp1a)            ' out
         angle1_oilgroove_2 = (angle_tmp0 + angle_tmp1b)            ' in

      '--- oilgroove arc2 range  out 0-1 and inner 0-2
         angle2_oilgroove_0 = (angle_tmp0 - angle_tmp2a + pi)       ' out
         angle2_oilgroove_1 = (angle_tmp0 - angle_tmp2b + pi)       ' in

   ElseIf ((pi / 2 <= angle_tmp0) And (angle_tmp0 < pi)) Then         ' PI/2 =< angle <PI
      '--- oilgroove arc1 range  out 0-1 and inner 0-2
         angle1_oilgroove_1 = (angle_tmp0 - angle_tmp1a)
         angle1_oilgroove_2 = (angle_tmp0 - angle_tmp1b)

      '--- oilgroove arc2 range  out 0-1 and inner 0-2
         angle2_oilgroove_0 = (angle_tmp0 + angle_tmp2a - pi)
         angle2_oilgroove_1 = (angle_tmp0 + angle_tmp2b - pi)

   ElseIf ((-pi < angle_tmp0) And (angle_tmp0 < -pi / 2)) Then       ' -PI< angle <-PI/2
      '--- oilgroove arc1 range  out 0-1 and inner 0-2
         angle1_oilgroove_1 = (angle_tmp0 - angle_tmp1a)
         angle1_oilgroove_2 = (angle_tmp0 - angle_tmp1b)

      '--- oilgroove arc2 range  out 0-1 and inner 0-2
         angle2_oilgroove_0 = (angle_tmp0 + angle_tmp2a + pi)
         angle2_oilgroove_1 = (angle_tmp0 + angle_tmp2b + pi)

   ElseIf ((-pi / 2 <= angle_tmp0) And (angle_tmp0 < 0)) Then        ' -PI/2 =< angle < 0
      '--- oilgroove arc1 range  out 0-2 and inner 1-2
         angle1_oilgroove_1 = (angle_tmp0 + angle_tmp1a)
         angle1_oilgroove_2 = (angle_tmp0 + angle_tmp1b)

      '--- oilgroove arc2 range  out 0-2 and inner 1-2
         angle2_oilgroove_0 = (angle_tmp0 - angle_tmp2a + pi)
         angle2_oilgroove_1 = (angle_tmp0 - angle_tmp2b + pi)

   End If

   '-----------
        q_tmp(11) = angle1_oilgroove_0        ' oil groove outer arc1 : start
        q_tmp(12) = angle1_oilgroove_1        ' oil groove outer arc1 :  connection point
         q_tmp(1) = angle1_oilgroove_0         ' oil groove inner arc1 : start
         q_tmp(2) = angle1_oilgroove_2         ' oil groove inner arc1 :   connection point

  '[PW-4]-----------
  '  oilgroove arc1 -> tp cross Point from start of arc1
        Phi_00 = angle1_oilgroove_0     ' =q_tmp(11), oil groove outer arc1 : start
         Phi_2 = angle1_oilgroove_1     ' =q_tmp(12), oil groove outer arc1 : connection point

         Phi_0 = angle1_oilgroove_0      ' =q_tmp(1) , oil groove inner arc1 : start
         Phi_1 = angle1_oilgroove_2      ' =q_tmp(2) , oil groove inner arc1 :  connection point


        Call Get_Gravity_Center_Oilgroove_PW_1(Phi_0, Phi_1, Phi_00, Phi_2)

            Area_tmp0 = Area_tmp
            xg_a_tmp0 = xg_a_tmp    'on FS xy-cordinate
            yg_a_tmp0 = yg_a_tmp    'on FS xy-cordinate

          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "PW-arc1_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(8, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(9, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

'GoTo Label_Gravity_center_end

   '[PW-5]-----------13
   '  oilgroove arc2 -> tp cross Point from start of arc1

        q_tmp(13) = angle2_oilgroove_0     ' oil groove outer arc2 : Connection point to end point
        q_tmp(14) = angle2_oilgroove_2     '
         q_tmp(5) = angle2_oilgroove_1     ' oil groove inner arc2 : Connection point to end point
         q_tmp(9) = angle2_oilgroove_2     ' ? q_tmp(6) = angle2_oilgroove_2

        Phi_00 = angle2_oilgroove_0        ' =q_tmp(13) oil groove outer arc2 : Connection point to end point
         Phi_2 = angle2_oilgroove_2        ' =q_tmp(14)
         Phi_0 = angle2_oilgroove_1        ' =q_tmp(5) oil groove inner arc2 : Connection point to end point
         Phi_1 = angle2_oilgroove_2        ' =q_tmp(9)

        Call Get_Gravity_Center_Oilgroove_PW_2(Phi_0, Phi_1, Phi_00, Phi_2)
            Sg2_b(7, I) = Area_tmp + Area_tmp0
            xg2_b(7, I) = (xg_a_tmp * Area_tmp + xg_a_tmp0 * Area_tmp0) / Sg2_b(7, I)
            yg2_b(7, I) = (yg_a_tmp * Area_tmp + yg_a_tmp0 * Area_tmp0) / Sg2_b(7, I)

          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "PW-arc2_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(10, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(11, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

        Debug.Print "PW Sg2_b(7, i) ="; Format(Sg2_b(7, I), "####.####"); Tab(2); _
                              "xg="; Format(xg2_b(7, I), "####.####"); "  "; _
                              "yg="; Format(yg2_b(7, I), "####.####")

'GoTo Label_Gravity_center_end      ' 20180514 check OK        (Contain offset dx,dy)


'-----------
'[Ml : area of between OS_sealring and Plate ]        Plate diameter = 125.00
'-----------

  '[ML-1]-----------
        ' inner arc1  (=seal Ring)    'on FS xy-cordinate
            Phi_0 = 0
            Phi_1 = 2 * pi

        ' outer arc1 (=OS_plate)      'on FS xy-cordinate
            Phi_00 = 0
            Phi_2 = 2 * pi

        Call Get_Gravity_Center_OS_plate_seal(Phi_0, Phi_1, Phi_00, Phi_2, the_1)

            Sg2_b(8, I) = Area_tmp
            xg2_b(8, I) = xg_a_tmp
            yg2_b(8, I) = yg_a_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "ML-1_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(12, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(13, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

  '[ML-2]-----------
  ' area of inner OS_sealring ]        Plate diameter = 125.00

         Area_tmp = (OS_seal) ^ 2 * pi / 4
         xg_a_tmp = 0                 ' Ro * Cos(the_1)
         yg_a_tmp = 0                 ' Ro * Sin(the_1)

            Sg2_b(9, I) = Area_tmp
            xg2_b(9, I) = xg_a_tmp
            yg2_b(9, I) = yg_a_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
'        Curve_name = "ML-2_":    ' DataSheetName = "DataSheet_2"
'        Call Paste_curve_data_Num(12, Curve_name & "in", x_in(), y_in(), DataSheetName)
'        Call Paste_curve_data_Num(13, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

'  '[Ml-]-----------
'  '  area of inner OS_sealring ]        Plate diameter = 125.00
'
'         Area_tmp = (OS_dia) ^ 2 * pi / 4
'         xg_a_tmp = Ro * Cos(the_1)
'         yg_a_tmp = Ro * Sin(the_1)
'
'            Sg2_b(10, i) = Area_tmp
'            xg2_b(10, i) = xg_a_tmp
'            yg2_b(10, i) = yg_a_tmp


'GoTo Label_Gravity_center_end


   ' 軸回転範囲 4PI = index 366 , 2PI:i=183 , PI:i=90 ,(i=0,45,90,138,183)
   '  i = 0:    the_1 = the(i) - qq
   '  the_1 = (the(i) - 0) * 180 / pi     ' [deg] on FS xy-cordinate
   '  the_1 = (the(i) - qq) * 180 / pi    ' [deg] on FS Wrap xy-cordinate
   '  the_1 = the(i) - qq                 ' [rad] on FS Wrap xy-cordinate
   '  div_n = 180

'-----------
'[tg10 : Between OS_Plate and oil groove]
'-----------

  '<tg10-1_1> start point of arc1 : outer oli groove 'on FS xy-cordinate
      x_tmp(11) = r1_oilgroove * Cos(angle1_oilgroove_0) + x1_oilgroove_c
      y_tmp(11) = r1_oilgroove * Sin(angle1_oilgroove_0) + y1_oilgroove_c
      q_tmp(11) = angle1_oilgroove_0

    ' end-connection point of oil groove's arc1 and arc2
      x_tmp(12) = r1_oilgroove * Cos(angle1_oilgroove_1) + x1_oilgroove_c
      y_tmp(12) = r1_oilgroove * Sin(angle1_oilgroove_1) + y1_oilgroove_c
      q_tmp(12) = angle1_oilgroove_1

  '<tg10-1_2> start point of arc1 OS Plate
        x1_tmp = x_tmp(11)     'on FS xy-cordinate
        y1_tmp = y_tmp(11)     'on FS xy-cordinate
        r1_tmp = OS_dia / 2
        x1c_tmp = Ro * Cos(the_1)
        y1c_tmp = Ro * Sin(the_1)

      'on FS xy-cordinate
        Call Get_CrossPoint_arc_on_line_xy(x1_tmp, y1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###

        x_tmp(31) = x1_tmp                           'on FS xy-cordinate
        y_tmp(31) = y1_tmp
        q_tmp(31) = Atan2((y_tmp(31) - y1c_tmp), (x_tmp(31) - x1c_tmp))

    '-- OS Plate arc1 : end connection point of arc1 OS Plate >
'        q1_tmp = Atan2((y_tmp(12)), (x_tmp(12)))     'on FS xy-cordinate
        x1_tmp = x_tmp(12)     'on FS xy-cordinate
        y1_tmp = y_tmp(12)     'on FS xy-cordinate
        r1_tmp = OS_dia / 2
        x1c_tmp = Ro * Cos(the_1)
        y1c_tmp = Ro * Sin(the_1)

      'on FS xy-cordinate
        Call Get_CrossPoint_arc_on_line_xy(x1_tmp, y1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###

        x_tmp(32) = x1_tmp           ' ==> tg6     'on FS xy-cordinate
        y_tmp(32) = y1_tmp
        q_tmp(32) = Atan2((y_tmp(32) - y1c_tmp), (x_tmp(32) - x1c_tmp))

    '-- OS Plate arc1 : start angle on FS xy-cordinate
         If q_tmp(31) > q_tmp(32) Then
             q_tmp(31) = Atan3((y_tmp(31) - y1c_tmp), (x_tmp(31) - x1c_tmp))
         End If
             angle1_OS_Plate_0 = q_tmp(31)
             angle1_OS_Plate_2 = q_tmp(32)

  '<tg10-1_3>  arc1
        Phi_00 = angle1_OS_Plate_0      ' =q_tmp(31),  outer arc1 : OS Plate  'on FS xy-cordinate
        Phi_2 = angle1_OS_Plate_2       ' =q_tmp(32),
        Phi_0 = angle1_oilgroove_0      ' =q_tmp(11),  inner arc1 : inner groove (start to Connection point)
        Phi_1 = angle1_oilgroove_2      ' =q_tmp(12),

        Call Get_Gravity_Center_Oilgroove_OS_tg10_1(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
            Area_tmp0 = Area_tmp
            xg_a_tmp0 = xg_a_tmp
            yg_a_tmp0 = yg_a_tmp

      '--Dara Paste -
        Curve_name = "tg10-1_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(14, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(15, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop


'GoTo Label_Gravity_center_end

   '<tg10-2_1> connection point of oli groove's right arc1 and left arc2
      x_tmp(13) = r2_oilgroove * Cos(angle2_oilgroove_0) + x2_oilgroove_c
      y_tmp(13) = r2_oilgroove * Sin(angle2_oilgroove_0) + y2_oilgroove_c
      q_tmp(13) = angle2_oilgroove_0

   ' end point of arc2 oli groove
      x_tmp(14) = r2_oilgroove * Cos(angle2_oilgroove_2) + x2_oilgroove_c
      y_tmp(14) = r2_oilgroove * Sin(angle2_oilgroove_2) + y2_oilgroove_c
      q_tmp(14) = angle2_oilgroove_2

   '<tg10-2_2> OS Plate arc2 : start angle on 0-xy(FS's)
      q_tmp(33) = Atan2((y_tmp(13) - y1c_tmp), (x_tmp(13) - x1c_tmp))
      q_tmp(34) = Atan2((y_tmp(14) - y1c_tmp), (x_tmp(14) - x1c_tmp))

         If q_tmp(33) > q_tmp(34) Then
               q_tmp(33) = Atan3((y_tmp(13) - y1c_tmp), (x_tmp(13) - x1c_tmp))
         End If

        angle1_OS_Plate_0 = q_tmp(33)
        angle1_OS_Plate_2 = q_tmp(34)

   '<tg10-2_3> left arc2

        Phi_00 = angle1_OS_Plate_0        ' =q_tmp(33) outer arc2 : OS Plate
           Phi_2 = angle1_OS_Plate_2      ' =q_tmp(34)
        Phi_0 = angle2_oilgroove_0        ' =q_tmp(13) inner arc2 : outer groove : Connection point to end point
           Phi_1 = angle2_oilgroove_2     ' =q_tmp(14)

        Call Get_Gravity_Center_Oilgroove_OS_tg10_2(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
            '    x_tmp(33),y_tmp(33)     ' start Point of outer arc ;  OS_Plate
            '    x_tmp(34),y_tmp(34)     ' end Point of outer arc ;  OS_Plate

            Sg_f(10, I) = Area_tmp + Area_tmp0
            xg_f(10, I) = (xg_a_tmp * Area_tmp + xg_a_tmp0 * Area_tmp0) / Sg_f(10, I)
            yg_f(10, I) = (yg_a_tmp * Area_tmp + yg_a_tmp0 * Area_tmp0) / Sg_f(10, I)

          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg10-2_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(16, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(17, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

        Debug.Print "Sg_f(10, i) ="; Format(Sg_f(10, I), "####.####"); Tab(2); _
                              "xg="; Format(xg_f(10, I), "####.####"); "  "; _
                              "yg="; Format(yg_f(10, I), "####.####")

'GoTo Label_Gravity_center_end    '-- [tg10] Ok  2018/05/10  (Contain dx,dy : on FS xy-cordinate)



   ' 軸回転範囲(i=0,45,90,138,183)    0-4PI = index 366 , 2PI:i=183 , PI:i=90 ,
   '  i = 0:  the_1 = the(i) - qq
   '  the_1 = (the(i) - 0) * 180 / pi     ' [deg] on FS xy-cordinate
   '  the_1 = (the(i) - qq) * 180 / pi    ' [deg] on FS Wrap xy-cordinate
   '  the_1 = the(i) - qq                 ' [rad] on FS Wrap xy-cordinate
   '  div_n = 180

'-----------
'[tg5 : FS Wrap Tip area of Between Suction inlet and FS_in ]
'-----------
   '  outer arc is suction inlet innner arc  => Refer to [Pg(3)] : suction inlet
   '    outer arc1 :   start      Phi_00 =          (the(i)=0)  <== inlet arc start
   '                   start      Phi_00 =          (the(i)<>0) <== inlet arc start+ FS inner curve
   '                   end        Phi_2  =
   '  inner arc1,2 are FS_in curve
   '    inner arc1 :   start      Phi_0  =  Phi_c_fi(3, i)  <== FS-OS contact point angle
   '                   end        Phi_1  =                  <== end angle of inlet arc
   '-----------

    '<tg5-1> outer Arc ; start and end point (= inlet-innnerRfo arc and  FS outer)

      '<outer arc1 start>
        'inlet-innner Rfo arc
        '  Cross point P25(x_tmp(25),y_tmp(25)) => ' see suction3-area3: inlet arc
           x_tmp(25) = x_tmp(25)
           y_tmp(25) = y_tmp(25)
           q_tmp(25) = q_tmp(25)

      '<outer arc end>
         'inlet-innner Rfo arc
         '  Cross point P22(x_tmp(22),y_tmp(22)) => ' see suction3-area3: inlet arc
            x_tmp(22) = x_tmp(22)
            y_tmp(22) = y_tmp(22)
            q_tmp(22) = q_tmp(22)

   '<tg5-2> inner Arc ; start and end point (=FS innner)
            tmp_q = Phi_c_fi(3, 0)
'            wrap_n = Int((tmp_q - qq) / (2 * pi))

      '<inner arc1 start>  ' Formura (8) fi         'on FS xy-cordinate
          Call Wp_xyfi(tmp_q)
            x_tmp(26) = Wp_xfi          ' x_tmp(26) = Formura (8) fi (tmp_q)
            y_tmp(26) = Wp_yfi          ' y_tmp(26) = Formura (8) fi (tmp_q)

'            x_tmp(26) = a * tmp_q ^ k * Cos(tmp_q - qq) _
'                                 + g1 * Cos(tmp_q - qq - Atn(k / tmp_q)) + dx
'            y_tmp(26) = a * tmp_q ^ k * Sin(tmp_q - qq) _
'                                 + g1 * Sin(tmp_q - qq - Atn(k / tmp_q)) + dy
            q_tmp(26) = tmp_q

       '<inner arc1 end>    ' Formura (8) fi  : connection point of Suction inlet
           'on FS xy-cordinate
           '  Cross point P22(x_tmp(22),y_tmp(22)) => see Pg(3) suction inlet

            tmp_q = Atan2((y_tmp(22) - dy), (x_tmp(22) - dx))
            wrap_n = Int((Phi_c_fi(3, 0) - tmp_q) / (2 * pi))

            tmp_q = tmp_q + 2 * pi * wrap_n + qq

          Call Wp_xyfi(tmp_q)
            x_tmp(23) = Wp_xfi          ' x_tmp(3) = Formura (8) fi (tmp_q)
            y_tmp(23) = Wp_yfi          ' y_tmp(3) = Formura (8) fi (tmp_q)

'            x_tmp(23) = a * tmp_q ^ k * Cos(tmp_q - qq) _
'                                 + g1 * Cos(tmp_q - qq - Atn(k / tmp_q)) + dx
'            y_tmp(23) = a * tmp_q ^ k * Sin(tmp_q - qq) _
'                                 + g1 * Sin(tmp_q - qq - Atn(k / tmp_q)) + dy
            q_tmp(23) = tmp_q


   '<tg5-3>
        ' inner arc1  (=FS innner)            'on FS xy-cordinate :  Wrap angle ramda
            Phi_0 = q_tmp(26)
            Phi_1 = q_tmp(23)

        ' outer arc1 (=inlet-innnerRfo arc)   'on FS xy-cordinate :  Wrap angle ramda
            Phi_00 = q_tmp(25)
            Phi_2 = q_tmp(22)

        Call Get_Gravity_Center_FS_wrapEnd_tg5(Phi_0, Phi_1, Phi_00, Phi_2, the_1)

            'on FS xy-cordinate
                Area_tmp0 = Area_tmp
                xg_a_tmp0 = xg_a_tmp
                yg_a_tmp0 = yg_a_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg5_1":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(18, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(19, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

'   GoTo Label_Gravity_center_end

    '<tg5-4>
      ' 1) Suction inlet : arc and Involute    Rotation angle: the_1 = Point(23) to 360deg
      ' 2) Suction inlet : arc only            Rotation angle: the_1 = 0 to Point(23)

         tmp_q = the_c(I) * 180 / pi

     If (0 < the_c(I)) And (the_c(I) <= (2 * pi - (q_tmp(23) - q_tmp(26)))) Then
'     If (0 < the_c(i)) And (the_c(i) <= ( pi) Then

        'inner arc2  (=FS innner)
          tmp_q = Phi_c_fi(3, I)

         '--<inner arc2 start>  ' Formura (8) fi      'on FS xy-cordinate
          Call Wp_xyfi(tmp_q)
            x_tmp(43) = Wp_xfi          ' x_tmp(43) = Formura (8) fi (tmp_q)
            y_tmp(43) = Wp_yfi          ' y_tmp(43) = Formura (8) fi (tmp_q)

'          x_tmp(43) = a * tmp_q ^ k * Cos(tmp_q - qq) _
'                               + g1 * Cos(tmp_q - qq - Atn(k / tmp_q)) + dx
'          y_tmp(43) = a * tmp_q ^ k * Sin(tmp_q - qq) _
'                               + g1 * Sin(tmp_q - qq - Atn(k / tmp_q)) + dy
          q_tmp(43) = tmp_q

            Phi_0 = q_tmp(43)
            Phi_1 = q_tmp(26)

        'outer arc2 (=FS outer)
          tmp_q = Phi_c_fi(2, I)

         '--<outer arc2 start>  ' Formura (7) fo       'on FS xy-cordinate
          x_tmp(42) = -a * tmp_q ^ k * Cos(tmp_q - qq) _
                               + g1 * Cos(tmp_q - qq - Atn(k / tmp_q)) + dx
          y_tmp(42) = -a * tmp_q ^ k * Sin(tmp_q - qq) _
                               + g1 * Sin(tmp_q - qq - Atn(k / tmp_q)) + dy
          q_tmp(42) = tmp_q

            Phi_00 = Phi_c_fi(2, I)    ' q_tmp(42)
            Phi_2 = Phi_c_fi(2, 0)     ' q_tmp(25)

        Call Get_Gravity_Center_FS_wrapEnd_tg5_2(Phi_0, Phi_1, Phi_00, Phi_2, the_1)

        'on FS xy-cordinate
            Sg_f(5, I) = Area_tmp + Area_tmp0
            xg_f(5, I) = (xg_a_tmp * Area_tmp + xg_a_tmp0 * Area_tmp0) / Sg_f(5, I)
            yg_f(5, I) = (yg_a_tmp * Area_tmp + yg_a_tmp0 * Area_tmp0) / Sg_f(5, I)
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg5_2":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(18, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(19, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

      Else

        Call Get_Gravity_Center_FS_wrapEnd_tg5_3(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
                Sg_f(5, I) = Area_tmp0
                xg_f(5, I) = xg_a_tmp0
                yg_f(5, I) = yg_a_tmp0
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg5_3":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(18, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(19, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

      End If

        Debug.Print "Sg_f(5, i) ="; Format(Sg_f(5, I), "####.####"); Tab(2); _
                              "xg="; Format(xg_f(5, I), "####.####"); "  "; _
                              "yg="; Format(yg_f(5, I), "####.####")

'   GoTo Label_Gravity_center_end
     '-- [tg5] 0-180 Ok 180-270?  2018/05/14  (Contain dx,dy : on FS xy-cordinate)



   ' 軸回転範囲(i=0,45,90,138,183)    0-4PI = index 366 , 2PI:i=183 , PI:i=90 ,
   '  i = 0:  the_1 = the(i) - qq
   '  the_1 = (the(i) - 0) * 180 / pi     ' [deg] on FS xy-cordinate
   '  the_1 = (the(i) - qq) * 180 / pi    ' [deg] on FS Wrap xy-cordinate
   '  the_1 = the(i) - qq                 ' [rad] on FS Wrap xy-cordinate
   '  div_n = 180

'-----------
'[tg6] : FS area of Between OS Plate and Suction inlet ]
'-----------

   '<tg6-1> inner Arc ; start and end point (=FS innner)
         wrap_n = Int((FS_in_end - qq) / (2 * pi))

      '<inner arc end>  ' Formura (8) fi
        ' angle1_oilgroove_0  'on FS wrap xy-cordinate
         x_tmp(1) = (r1_oilgroove - t1_oilgroove) * Cos(angle1_oilgroove_0) + x1_oilgroove_c
         y_tmp(1) = (r1_oilgroove - t1_oilgroove) * Sin(angle1_oilgroove_0) + y1_oilgroove_c

      '<inner arc1 start>  ' Formura (8) fi
         angle1_tg_0 = Atan2((y_tmp(1) - dy), (x_tmp(1) - dx)) + qq    '+ 2 * pi * wrap_n
                   ' = -68.11 + 360*2 =-68.11+720 +105 = 756.89
         angle1_tg_0 = angle1_tg_0 + 2 * pi * Int((FS_in_end - angle1_tg_0) / (2 * pi))   '##

        Call Wp_xyfi(angle1_tg_0)
            x_tmp(3) = Wp_xfi          ' x_tmp(3) = Formura (8) fi (angle1_tg_0)
            y_tmp(3) = Wp_yfi          ' y_tmp(3) = Formura (8) fi (angle1_tg_0)

'           x_tmp(3) = x_tmp(3)                  ' see tg7
'           y_tmp(3) = y_tmp(3)

          'on FS wrap xy-cordinate
           q_tmp(3) = Atan2((y_tmp(3) - dy), (x_tmp(3) - dx)) + 2 * pi * wrap_n + qq

            If q_tmp(3) > (FS_in_end - qq) Then
                q_tmp(3) = q_tmp(3) - 2 * pi
            End If

      '<inner arc start>  ' Formura (8) fi
           x_tmp(23) = x_tmp(23)                ' see tg5
           y_tmp(23) = y_tmp(23)

        'on FS wrap xy-cordinate
           q_tmp(23) = Atan2((y_tmp(23) - dy), (x_tmp(23) - dx)) + 2 * pi * wrap_n + qq

            If q_tmp(23) > (FS_in_end - qq) Then
                q_tmp(23) = q_tmp(23) - 2 * pi
            End If

   '<tg6-2> <outer arc end>  ' OS Plate
           x_tmp(31) = x_tmp(31)    'on FS xy-cordinate   ' see tg10
           y_tmp(31) = y_tmp(31)    'on FS xy-cordinate

          'on FS wrap xy-cordinate
           q_tmp(31) = Atan2((y_tmp(31) - Ro * Sin(the_1)), (x_tmp(31) - Ro * Cos(the_1)))

      '<outer arc start>   ' OS Plate
'           q1_tmp = Atan2((y_tmp(21)), (x_tmp(21)))     'on FS xy-cordinate   ' see tg5 or Pg(3)
            x1_tmp = x_tmp(21)     'on FS xy-cordinate
            y1_tmp = y_tmp(21)     'on FS xy-cordinate
            r1_tmp = OS_dia / 2
            x1c_tmp = Ro * Cos(the_1)
            y1c_tmp = Ro * Sin(the_1)

              'on FS xy-cordinate
        Call Get_CrossPoint_arc_on_line_xy(x1_tmp, y1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###

           x_tmp(35) = x1_tmp           ' ==> tg9     'on FS xy-cordinate
           y_tmp(35) = y1_tmp

          'on OS Plate xy-cordinate
             q_tmp(35) = Atan2((y_tmp(35) - Ro * Sin(the_1)), (x_tmp(35) - Ro * Cos(the_1)))

      ' angle :
            Phi_0 = q_tmp(23)          ' start of inner angle    on FS wrap xy-cordinate
            Phi_1 = q_tmp(3)           ' end of inner angle      on FS wrap xy-cordinate

            Phi_00 = q_tmp(35)         ' start of outer angle    on OS Plate xy-cordinate
            Phi_2 = q_tmp(31)          ' end of outer angle      on OS Plate xy-cordinate

          Call Get_Gravity_Center_FS_OS_tg6(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
                Sg_f(6, I) = Area_tmp
                xg_f(6, I) = xg_a_tmp
                yg_f(6, I) = yg_a_tmp
          Call change_Wrap_data_to_curve_xw(2, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(3, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg6_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(12, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(13, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

        ' Routin to draw OS plate outer circle

        '  Call Data_Strage_to_array_check_curve

        Debug.Print "Sg_f(6, i) ="; Format(Sg_f(6, I), "####.####"); Tab(2); _
                              "xg="; Format(xg_f(6, I), "####.####"); "  "; _
                              "yg="; Format(yg_f(6, I), "####.####")

'        Debug.Print "Data_Strage_to_array_check_curve"
        Debug.Print ""

'GoTo Label_Gravity_center_end


   ' 軸回転範囲 (i=0,45,90,138,183)   0, PI/2(i=45), PI(i=90),1.5PI(=138) 2PI(i=183), 4PI(i= 366)
   '  i = 138:    the_1 = the(i) - qq
   '  the_1 = (the(i) - 0) * 180 / pi     ' [deg] on FS xy-cordinate
   '  the_1 = (the(i) - qq) * 180 / pi    ' [deg] on FS Wrap xy-cordinate
   '  the_1 = the(i) - qq                 ' [rad] on FS Wrap xy-cordinate
   '  div_n = 180


'-----------
'[tg7] : [area of Between oil groove inner arc1 and FS_in ]
'-----------
   '  outer arc1,2 are oil groove innner arc   => Refer to [pw] : oil groove
   '    outer arc1 :   start       =  angle1_oilgroove_0     <== oilgroove start
   '                   connection  =  angle1_oilgroove_2
   '    outer arc2 :   start       =  angle2_oilgroove_1
   '                   end angle   =  angle2_oilgroove_2
   '---
   '  inner arc1,2 are FS_in curve
   '    inner arc1 :   start       =  angle1_tg_0            <== FS=in start
   '                   connection  =  angle1_tg_2
   '    inner arc2 :   start       =  angle2_tg_1
   '                   end         =  angle2_tg_2            <== Phi_c_fi(1, i)
   '-----------

'  tg7 inner arc
'   1) i= G[0] to G[2]    G[0] Conpression start  G[2] Connection point
'   2) i= G[2] to G[4]    G[2] Connection point    G[4] oil-groove end
'   3) i= G[4] to G[6]    G[4] oil-groove_end     G[6] oil-groove_start
'   4) i= G[6] to G[8]    G[6] oil-groove_start   G[8] Conpression start
'


'1) i= G[0] to G[2]    G[0] Conpression start  G[2] Connection point

  '<tg7-1-1>  start and end point : arc1 and arc2

    '<outer arc1 start point>
      ' angle1_oilgroove_0
       q_tmp(1) = angle1_oilgroove_0      ' setting data
       x_tmp(1) = (r1_oilgroove - t1_oilgroove) * Cos(angle1_oilgroove_0) + x1_oilgroove_c
       y_tmp(1) = (r1_oilgroove - t1_oilgroove) * Sin(angle1_oilgroove_0) + y1_oilgroove_c


    '<inner arc1 start point>  ' Formura (8) fi
        tmp_q = Atan2((y_tmp(1) - dy), (x_tmp(1) - dx)) + qq
                ' = -68.11 + 360*2 =-68.11+720 +105 = 756.89
        angle1_tg_0 = tmp_q + 2 * pi * Int((FS_in_end - tmp_q) / (2 * pi))   '##
          q_tmp(3) = angle1_tg_0

        Call Wp_xyfi(angle1_tg_0)
          x_tmp(3) = Wp_xfi          ' x_tmp(3) = Formura (8) fi (angle1_tg_0)
          y_tmp(3) = Wp_yfi          ' y_tmp(3) = Formura (8) fi (angle1_tg_0)


  '<tg7-1-2> end angle right arc1 : FS_in curve1

      '<outer arc1 end>   ' connection point of oil groove's right arc1 and left arc2
          ' angle1_oilgroove_1
          x_tmp(2) = (r1_oilgroove - t1_oilgroove) * Cos(angle1_oilgroove_2) + x1_oilgroove_c
          y_tmp(2) = (r1_oilgroove - t1_oilgroove) * Sin(angle1_oilgroove_2) + y1_oilgroove_c
          q_tmp(2) = angle1_oilgroove_2

      '<inner arc1 end>    right FS_in  : connection point angle on 0-xy(FS's)
          wrap_n = Int((FS_in_end - Atan2((y_tmp(2) - dy), (x_tmp(2) - dx))) / (2 * pi))   '##
          angle1_tg_2 = Atan2((y_tmp(2) - dy), (x_tmp(2) - dx)) + 2 * pi * wrap_n + qq
                    ' = 84.7 + 360*2 =84.7+720 +0 = 804.7

      '--- FS contact point is on arc1
        If (angle1_tg_0 < Phi_c_fi(1, I) And Phi_c_fi(1, I) <= angle1_tg_2) Then

            q_rfp(4) = Phi_c_fi(1, I)

            Call Wp_xyfi(q_rfp(4))    ' FS contact point is on arc1
              x_rfp(4) = Wp_xfi          ' x_tmp(4) = Formura (8) fi (angle1_tg_2)
              y_rfp(4) = Wp_yfi          ' y_tmp(4) = Formura (8) fi (angle1_tg_2)

          ' cross point between FS_in and inner oil groove arc
'              q1_tmp = Atan2((y_tmp(4)), (x_tmp(4)))       'see x_tmp(9) of PW   'on FS xy-cordinate
              x1_tmp = x_rfp(4)     'on FS xy-cordinate
              y1_tmp = y_rfp(4)     'on FS xy-cordinate
              r1_tmp = r1_oilgroove - t1_oilgroove
              x1c_tmp = x1_oilgroove_c
              y1c_tmp = y1_oilgroove_c

            '**** =[-b-Root(b^2-4ac)]/(2a)
             'Call Get_CrossPoint_arc_on_line_2(q1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###
            '**** =[-b+Root(b^2-4ac)]/(2a)
             'Call Get_CrossPoint_arc_on_line(q1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)   '###

            Call Get_CrossPoint_arc_on_line_xy(x1_tmp, y1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###

              x_rfp(2) = x1_tmp                                   ' on FS xy-cordinate
              y_rfp(2) = y1_tmp
              q_rfp(2) = Atan2((y_rfp(2) - y1_oilgroove_c), (x_rfp(2) - x1_oilgroove_c))

              If q_rfp(2) > angle1_oilgroove_2 Then
                 q_rfp(2) = q_rfp(2) - 2 * pi
              End If


        ElseIf (angle1_tg_2 < Phi_c_fi(1, I) And Phi_c_fi(1, I) <= Phi_c_fi(1, 0)) Then   ' Sg[tg7] = [1-Max]
            q_rfp(4) = angle1_tg_2

            Call Wp_xyfi(q_rfp(4))    ' FS contact point is on arc1
              x_rfp(4) = Wp_xfi          ' x_tmp(4) = Formura (8) fi (angle1_tg_2)
              y_rfp(4) = Wp_yfi          ' y_tmp(4) = Formura (8) fi (angle1_tg_2)

              q_rfp(2) = angle1_oilgroove_2

        Else
        ' ElseIf (angle1_tg_0 >= Phi_c_fi(1, i) And Phi_c_fi(1, i) > Phi_c_fi(1, 0)) Then  ' Sg[tg7] = 0
        ' ElseIf Phi_c_fi(1, i) = angle1_tg_0 Then    ' Sg[tg7] = 0

            q_rfp(4) = q_tmp(3)             ' q_tmp(3) = angle1_tg_0
            q_rfp(2) = q_tmp(1)             ' q_tmp(2) = angle1_oilgroove_0

        End If

   '<tg7-1-3> angle range of arc1s : oil groove innner arc and FS_in curve1

            Phi_00 = q_tmp(1)     '= angle1_oilgroove_0   : outer arc1 start : innner oil groove
            Phi_2 = q_rfp(2)      '=moving or angle1_oilgroove_2   : outer arc1 end
                                  '     : groove Connection point /or wrap contact angle
            Phi_0 = q_tmp(3)      '= angle1_tg_0                   : inner arc1 start : FS_in start
            Phi_1 = q_rfp(4)      '= moving Phi_c_fi(1, i) or angle1_tg_2          : inner arc1 end
                                  ' q_tmp(4): FS_in Connection point /or wrap contact angle

        Call Get_Gravity_Center_Oilgroove_FS_tg7_1(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
            Area_tmp0 = Area_tmp
            xg_a_tmp0 = xg_a_tmp
            yg_a_tmp0 = yg_a_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg7-1_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(14, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(15, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

'GoTo Label_Gravity_center_end

   '<tg7-2-1>  start point of arc2 : FS_in curve

        '<outer arc2 start>   ' connection point of oil grooves arc1 and left arc2
        '   angle2_oilgroove_1
            x_tmp(5) = (r1_oilgroove - t1_oilgroove) * Cos(angle2_oilgroove_1) + x1_oilgroove_c
            y_tmp(5) = (r1_oilgroove - t1_oilgroove) * Sin(angle2_oilgroove_1) + y1_oilgroove_c
            q_tmp(5) = angle2_oilgroove_1

        '<inner arc2 start>   FS_in  : connection point of oil grooves on 0-xy(FS's)
            angle2_tg_1 = angle1_tg_2   ' q_tmp(7)= q_tmp(4) Connection point of arc1 and arc2
            angle2_tg_2 = Phi_c_fi(1, 0)

            q_tmp(7) = angle2_tg_1
              Call Wp_xyfi(q_tmp(7))        ' q_tmp(4=7) = angle1_tg_2 FS wrap contact point
            x_tmp(7) = Wp_xfi           ' x_tmp(4) = Formura (8) fi (angle1_tg_2)
            y_tmp(7) = Wp_yfi           ' y_tmp(4) = Formura (8) fi (angle1_tg_2)

            q_tmp(8) = Phi_c_fi(1, 0)     ' q_tmp(8) = angle2_tg_2 = Phi_c_fi(1, 0)
              Call Wp_xyfi(q_tmp(8))        ' q_tmp(4=7) = angle1_tg_2 FS wrap contact point
            x_tmp(8) = Wp_xfi           ' x_tmp(4) = Formura (8) fi (angle1_tg_2)
            y_tmp(8) = Wp_yfi

            q_rfp(8) = Phi_c_fi(1, I)

      ' If (angle1_tg_2 < Phi_c_fi(1, i)) And (Phi_c_fi(1, i) <= angle2_tg_2) Then

      If (q_tmp(7) < q_rfp(8)) And (q_rfp(8) <= q_tmp(8)) Then

        '<inner arc2 end  :Moving point>   FS_in
            q_rfp(8) = Phi_c_fi(1, I)     ' wrap contact point
              Call Wp_xyfi(q_rfp(8))
            x_rfp(8) = Wp_xfi          ' x_tmp(4) = Formura (8) fi (tmp_q)
            y_rfp(8) = Wp_yfi          ' y_tmp(4) = Formura (8) fi (tmp_q)

        '<outer arc2 end :Moving point>  oil groove
            tmp_q = Atan2((y_rfp(8) - y2_oilgroove_c), (x_rfp(8) - x2_oilgroove_c))

            x_rfp(6) = (r2_oilgroove - t2_oilgroove) * Cos(tmp_q) + x2_oilgroove_c
            y_rfp(6) = (r2_oilgroove - t2_oilgroove) * Sin(tmp_q) + y2_oilgroove_c
            q_rfp(6) = tmp_q

         '<arc2 angle range>
            Phi_00 = q_tmp(5)     '= angle2_oilgroove_0 ' outer arc2 start : Connection point of oil groove
            Phi_2 = q_rfp(6)      '= Moving Point       ' outer arc2 end   : oil groove end
            Phi_0 = q_tmp(7)      '= angle2_tg_1        ' inner arc2 start : FS_in  Connection point
            Phi_1 = q_rfp(8)      '= Moving Point       ' inner arc2 end   : FS_in  end

      Else         ' Phi_c_fi(1, i) <= angle1_tg_2       '  S(tg7-2)=0

         '<arc2 angle range>
            Phi_00 = 0           '= q_tmp(5) outer arc2 start : Connection point of oil groove
            Phi_2 = 0            '= q_tmp(6) outer arc2 end   : oil groove end
            Phi_0 = 0            '= q_tmp(7) inner arc2 start : FS_in  Connection point
            Phi_1 = 0            '= q_tmp(8) inner arc2 end   : FS_in  end

      End If

        ' Sg7([1]+[2])
        Call Get_Gravity_Center_Oilgroove_FS_tg7_2(Phi_0, Phi_1, Phi_00, Phi_2, the_1)

            Sg_f(7, I) = Area_tmp + Area_tmp0
            xg_f(7, I) = (xg_a_tmp * Area_tmp + xg_a_tmp0 * Area_tmp0) / Sg_f(7, I)
            yg_f(7, I) = (yg_a_tmp * Area_tmp + yg_a_tmp0 * Area_tmp0) / Sg_f(7, I)

          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg7-2_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(16, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(17, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

        Debug.Print "Sg_f(7, i) ="; Format(Sg_f(7, I), "####.####"); Tab(2); _
                              "xg="; Format(xg_f(7, I), "####.####"); "  "; _
                              "yg="; Format(yg_f(7, I), "####.####")

'GoTo Label_Gravity_center_end     '---- tg7 Ok  2018/05/8 ,5/14     (Contain offset dx,dy)



   ' 軸回転範囲(i=0,45,90,138,183)    0-4PI = index 366 , 2PI:i=183 , PI:i=90 ,
   '  i = 0:  the_1 = the(i) - qq
   '  the_1 = (the(i) - 0) * 180 / pi     ' [deg] on FS xy-cordinate
   '  the_1 = (the(i) - qq) * 180 / pi    ' [deg] on FS Wrap xy-cordinate
   '  the_1 = the(i) - qq                 ' [rad] on FS Wrap xy-cordinate
   '  div_n = 180

'    ReDim xfi(div_n):   ReDim yfi(div_n)
'    ReDim xfo(div_n):   ReDim yfo(div_n)

'-----------
'[tg8 : FS area of Between oil-groove and Suction-Inlet arc (FS_in) ]
'-----------
    '[tg8-1]    angle1_tg_2  < Phi_c_fi(1, i) <= angle1_tg_0)     ' S(tg8[1]+[2]+[3])
    '[tg8-2]  Phi_c_fi(1, 0) < Phi_c_fi(1, i) <= angle1_tg_2)     ' S(tg8[2]+[3])
    '[tg8-3]  Phi_c_fi(1, 0) = Phi_c_fi(1, i)                     ' S(tg8[3])


   '<tg8-0> outer arc3 ; start and end point (= innner oil groove)

            q_tmp(26) = Phi_c_fi(3, 0)

      '<inner arc start>  ' Suction inlet inner arc

            q_tmp(8) = Phi_c_fi(1, 0)    'Formura (8) fi
          Call Wp_xyfi(q_tmp(8))         ' q_tmp(4=7) = angle1_tg_2 FS wrap contact point
            x_tmp(28) = Wp_xfi           ' x_tmp(8) = Formura (8) fi (angle2_tg_2)
            y_tmp(28) = Wp_yfi           ' y_tmp(8) = Formura (8) fi (angle2_tg_2)

          ' Suction inlet inner arc : Rfi
            q_tmp(28) = Atan2((y_tmp(28) - y_Rfi_c), (x_tmp(28) - x_Rfi_c))

      '<outer arc start>  ' innner oil groove
            tmp_q = Atan2((y_tmp(28) - y2_oilgroove_c), (x_tmp(28) - x2_oilgroove_c))

            x_tmp(27) = (r2_oilgroove - t2_oilgroove) * Cos(tmp_q) + x2_oilgroove_c
            y_tmp(27) = (r2_oilgroove - t2_oilgroove) * Sin(tmp_q) + y2_oilgroove_c
            q_tmp(27) = tmp_q

      '<outer arc end>   ' innner oil groove  ' see PW     'on FS xy-cordinate
            tmp_q = angle2_oilgroove_2 ' oil groove inner arc1 :   connection point

            x_tmp(9) = (r2_oilgroove - t2_oilgroove) * Cos(tmp_q) + x2_oilgroove_c
            y_tmp(9) = (r2_oilgroove - t2_oilgroove) * Sin(tmp_q) + y2_oilgroove_c
            q_tmp(9) = angle2_oilgroove_2

      '<inner arc end>  ' Suction inlet inner arc
          ' cross point between FS_in and Suction inlet arc
          ' q1_tmp = Atan2((y_tmp(9)), (x_tmp(9)))    'see x_tmp(9) of PW   'on FS xy-cordinate

            x1_tmp = x_tmp(9)     'on FS xy-cordinate
            y1_tmp = y_tmp(9)     'on FS xy-cordinate
             r1_tmp = r_Rfi_c
             x1c_tmp = x_Rfi_c
             y1c_tmp = y_Rfi_c

          'on FS xy-cordinate
            Call Get_CrossPoint_arc_on_line_xy(x1_tmp, y1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###
            ' Call Get_CrossPoint_arc_on_line_2(q1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)   '###

            x_tmp(24) = x1_tmp                                   ' on FS xy-cordinate
            y_tmp(24) = y1_tmp
            q_tmp(24) = Atan2((y_tmp(24) - y_Rfi_c), (x_tmp(24) - x_Rfi_c))

      '<arc3 angle range>
            Phi_0 = q_tmp(28)         ' start angle of inner arc    on FS xy-cordinate
            Phi_1 = q_tmp(24)         ' end angle of inner arc      on FS xy-cordinate

            Phi_00 = q_tmp(27)        ' start angle of outer arc    'on FS Wrap xy-cordinate
            Phi_2 = q_tmp(9)          ' end angle of outer arc      'on FS Wrap xy-cordinate

      ' S(tg8[3])
        Call Get_Gravity_Center_Oilgroove_FS_tg8_3(Phi_0, Phi_1, Phi_00, Phi_2, the_1)

            Area_tmp0 = Area_tmp
            xg_a_tmp0 = xg_a_tmp
            yg_a_tmp0 = yg_a_tmp

       '--Dara Paste -
        Curve_name = "tg8-3_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(20, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(21, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop


'GoTo Label_Gravity_center_end    '<tg8>


'  If (q_tmp(26) < Phi_c_fi(1, i) And Phi_c_fi(1, i) <= angle1_tg_2) Then

  If (angle1_tg_0 < Phi_c_fi(1, I) And Phi_c_fi(1, I) <= angle1_tg_2) Then

  ' FS contact point is on arc1
  '<tg8-1>    : S(tg8[1]+[2]+[3])  'FS contact point is on arc1

  '<tg8-1-1>
      '<inner arc1 start point>  FS_in
            q_rfp(4) = Phi_c_fi(1, I)

          Call Wp_xyfi(q_rfp(4))        ' FS contact point is on arc1
            x_rfp(4) = Wp_xfi           ' x_tmp(4) = Formura (8) fi ()
            y_rfp(4) = Wp_yfi           ' y_tmp(4) = Formura (8) fi ()

      '<outer arc1 start point>   ' connection point of oil groove's right arc1 and left arc2
        ' cross point between FS_in and inner oil groove arc
'            q1_tmp = Atan2((y_tmp(4)), (x_tmp(4)))       'see x_tmp(9) of PW   'on FS xy-cordinate
          x1_tmp = x_rfp(4)     'on FS xy-cordinate
          y1_tmp = y_rfp(4)     'on FS xy-cordinate
            r1_tmp = r1_oilgroove - t1_oilgroove
            x1c_tmp = x1_oilgroove_c
            y1c_tmp = y1_oilgroove_c

          'on FS xy-cordinate
            Call Get_CrossPoint_arc_on_line_xy(x1_tmp, y1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###
          ' Call Get_CrossPoint_arc_on_line_2(q1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)   '###

            x_rfp(2) = x1_tmp                                   ' on FS xy-cordinate
            y_rfp(2) = y1_tmp
            q_rfp(2) = Atan2((y_rfp(2) - y1_oilgroove_c), (x_rfp(2) - x1_oilgroove_c))

              If q_rfp(2) > angle1_oilgroove_2 Then
                 q_rfp(2) = q_rfp(2) - 2 * pi
              End If

      '<inner arc1 end point> FS_in
            q_tmp(4) = angle1_tg_2     ' q_tmp(7) = angle2_tg_1

          Call Wp_xyfi(q_tmp(4))       ' FS contact point is on arc1
            x_tmp(4) = Wp_xfi           ' x_tmp(4) = Formura (8) fi (angle1_tg_2)
            y_tmp(4) = Wp_yfi           ' y_tmp(4) = Formura (8) fi (angle1_tg_2)

      '<outer arc1 end point>   ' connection point of oil groove's right arc1 and left arc2
            x_tmp(2) = (r1_oilgroove - t1_oilgroove) * Cos(angle1_oilgroove_1) + x1_oilgroove_c
            y_tmp(2) = (r1_oilgroove - t1_oilgroove) * Sin(angle1_oilgroove_1) + y1_oilgroove_c
            q_tmp(2) = angle1_oilgroove_2

      '<arc1 angle range>
            Phi_00 = q_rfp(2)   ' Moving Point        ' outer arc1 start : innner groove /or wrap contact angle
            Phi_2 = q_tmp(2)    ' angle1_oilgroove_2  '    end   : innner oil groove
            Phi_0 = q_rfp(4)    ' Moving Point        ' inner arc1 start : FS_in start
            Phi_1 = q_tmp(4)    ' angle1_tg_2         '    end   : FS_in           /or wrap contact angle

        Call Get_Gravity_Center_Oilgroove_FS_tg8_1(Phi_0, Phi_1, Phi_00, Phi_2, the_1)

       '--Dara Paste -
        Curve_name = "tg8-1-1_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(20, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(21, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

      ' Sg8[1]+[3]
            Sg_f(8, I) = Area_tmp + Area_tmp0
            xg_f(8, I) = (xg_a_tmp * Area_tmp + xg_a_tmp0 * Area_tmp0) / Sg_f(8, I)
            yg_f(8, I) = (yg_a_tmp * Area_tmp + yg_a_tmp0 * Area_tmp0) / Sg_f(8, I)

            Area_tmp0 = Sg_f(8, I)
            xg_a_tmp0 = xg_f(8, I)
            xg_a_tmp0 = yg_f(8, I)


   '<tg8-1-2>

      '<outer arc2 start & end  point> : innner oil groove
            q_tmp(5) = angle2_oilgroove_1   ' Connection point of arc1 and arc2
      '<outer arc2 end point>
            q_tmp(6) = q_tmp(27)            '

      '<inner arc2 start point> : FSin
            q_tmp(7) = angle2_tg_1          ' Connection point of arc1 and arc2
      '<inner arc2 end point>
            q_tmp(8) = q_tmp(8)            'Phi_c_fi(1, 0) = Compression Start point of A chamber

      '<arc2 angle range>
            Phi_00 = q_tmp(5)   ' outer arc1 start : groove Connection point /or wrap contact angle
            Phi_2 = q_tmp(6)    '    end   : innner oil groove
            Phi_0 = q_tmp(7)    ' inner arc1 start : FS_in start
            Phi_1 = q_tmp(8)    '    end   : FS_in Connection point /or wrap contact angle

        Call Get_Gravity_Center_Oilgroove_FS_tg8_2(Phi_0, Phi_1, Phi_00, Phi_2, the_1)

        ' Sg8([1]+[3]) + Sg8([2])
            Sg_f(8, I) = Area_tmp + Area_tmp0
            xg_f(8, I) = (xg_a_tmp * Area_tmp + xg_a_tmp0 * Area_tmp0) / Sg_f(8, I)
            yg_f(8, I) = (yg_a_tmp * Area_tmp + yg_a_tmp0 * Area_tmp0) / Sg_f(8, I)
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg8-1-2_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(22, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(23, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

  ElseIf (angle1_tg_2 < Phi_c_fi(1, I) And Phi_c_fi(1, I) < q_tmp(8)) Then    'FS contact point is on arc2

    '<tg8-2>    : S(tg8[2]+[3])


         '<inner arc2 start : Moving point>   FS_in
            q_rfp(8) = Phi_c_fi(1, I)     ' wrap contact point
              Call Wp_xyfi(q_rfp(8))
            x_rfp(8) = Wp_xfi          ' x_tmp(4) = Formura (8) fi (tmp_q)
            y_rfp(8) = Wp_yfi          ' y_tmp(4) = Formura (8) fi (tmp_q)

        '<outer arc2 start : Moving point>  oil groove
            tmp_q = Atan2((y_rfp(8) - y2_oilgroove_c), (x_rfp(8) - x2_oilgroove_c))

            x_rfp(6) = (r2_oilgroove - t2_oilgroove) * Cos(tmp_q) + x2_oilgroove_c
            y_rfp(6) = (r2_oilgroove - t2_oilgroove) * Sin(tmp_q) + y2_oilgroove_c
            q_rfp(6) = tmp_q

        '<outer arc2 end point>
              q_tmp(6) = q_tmp(27)      ' (angle2_oilgroove_2)

        '<inner arc2 end point>
              q_tmp(8) = q_tmp(8)             '

      '<arc2 angle range>
            Phi_00 = q_rfp(6)   ' outer arc1 start  : innner oil groove
            Phi_2 = q_tmp(6)    '    end           : innner oil groove
            Phi_0 = q_rfp(8)    ' inner arc1 start  : FS_in Connection point /or wrap contact angle
            Phi_1 = q_tmp(8)    '    end           : FS_in end

        Call Get_Gravity_Center_Oilgroove_FS_tg8_2(Phi_0, Phi_1, Phi_00, Phi_2, the_1)

        '  S(tg8[2]+[3])
            Sg_f(8, I) = Area_tmp + Area_tmp0
            xg_f(8, I) = (xg_a_tmp * Area_tmp + xg_a_tmp0 * Area_tmp0) / Sg_f(8, I)
            yg_f(8, I) = (yg_a_tmp * Area_tmp + yg_a_tmp0 * Area_tmp0) / Sg_f(8, I)
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg8-2_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(22, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(23, Curve_name & "out", x_out(), y_out(), DataSheetName)
   Stop

  ElseIf Phi_c_fi(1, I) = q_tmp(8) Then     '  If the(i) = the(0)  q_tmp(28) = Phi_c_fi(1, 0)

        '  S(tg8[3])
          Sg_f(8, I) = Area_tmp0
          xg_f(8, I) = xg_a_tmp0
          yg_f(8, I) = yg_a_tmp0

  Else
        Stop

  End If

        Debug.Print "Sg_f(8, i) ="; Format(Area_tmp, "####.####"); Tab(2); _
                              "xg="; Format(xg_f(8, I), "####.####"); "  "; _
                              "yg="; Format(yg_f(8, I), "####.####")


'GoTo Label_Gravity_center_end    '<tg8>


   ' 軸回転範囲(i=0,45,90,138,183)    0-4PI = index 366 , 2PI:i=183 , PI:i=90 ,
   '  i = 45:  the_1 = the(i) - qq
   '  the_1 = (the(i) - 0) * 180 / pi     ' [deg] on FS xy-cordinate
   '  the_1 = (the(i) - qq) * 180 / pi    ' [deg] on FS Wrap xy-cordinate
   '  the_1 = the(i) - qq                 ' [rad] on FS Wrap xy-cordinate
   '  div_n = 180

'-----------
'[tg9 : FS area of Between oil groove and Suction inlet ]
'-----------
   '<tg9> outer Arc ; start and end point (= OS plate)

      '<outer arc start>  ' OS Plate
           x_tmp(34) = x_tmp(34)          ' see tg10     'on FS xy-cordinate
           y_tmp(34) = y_tmp(34)
           q_tmp(34) = Atan2((y_tmp(34) - Ro * Sin(the_1)), (x_tmp(34) - Ro * Cos(the_1)))

      '<outer arc end>   ' OS Plate
           x_tmp(35) = x_tmp(35)          ' see tg6     'on FS xy-cordinate
           y_tmp(35) = y_tmp(35)
           q_tmp(35) = Atan2((y_tmp(35) - Ro * Sin(the_1)), (x_tmp(35) - Ro * Cos(the_1)))

   '<tg9> inner arc ; start and end point (= suction inlet outer-arc)

      '<inner arc start>  ' inlet outer arc
'           q1_tmp = Atan2(y_tmp(34), x_tmp(34)) * 180 / pi
'           q1_tmp = Atan2(y_tmp(34) - y_Rfi_c, x_tmp(34) - x_Rfi_c) * 180 / pi

           ' q1_tmp = Atan2((y_tmp(34)), (x_tmp(34)))     'on FS xy-cordinate
          x1_tmp = x_tmp(34)     'on FS xy-cordinate
          y1_tmp = y_tmp(34)     'on FS xy-cordinate
           r1_tmp = r_Rfi_c
           x1c_tmp = x_Rfi_c
           y1c_tmp = y_Rfi_c

      'on FS xy-cordinate
        Call Get_CrossPoint_arc_on_line_xy(x1_tmp, y1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###
      ' Call Get_CrossPoint_arc_on_line_2(q1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)   '###

           x_tmp(24) = x1_tmp           ' ==> tg6     'on FS xy-cordinate
           y_tmp(24) = y1_tmp
           q_tmp(24) = Atan2(y_tmp(24) - y_Rfi_c, x_tmp(24) - x_Rfi_c)   'center of Rfi on FS xy-cordinate

      '<inner arc end>    '  inlet outer arc
           ' q1_tmp = Atan2((y_tmp(35)), (x_tmp(35)))     'on FS xy-cordinate
          x1_tmp = x_tmp(35)     'on FS xy-cordinate
          y1_tmp = y_tmp(35)     'on FS xy-cordinate
           r1_tmp = r_Rfi_c
           x1c_tmp = x_Rfi_c
           y1c_tmp = y_Rfi_c

      'on FS xy-cordinate
        Call Get_CrossPoint_arc_on_line_xy(x1_tmp, y1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)  '###
      ' Call Get_CrossPoint_arc_on_line(q1_tmp, r1_tmp, x1c_tmp, y1c_tmp, the_1)   '###

           x_tmp(21) = x1_tmp           ' ==> tg6     'on FS xy-cordinate
           y_tmp(21) = y1_tmp
           q_tmp(21) = Atan2(y_tmp(21) - y_Rfi_c, x_tmp(21) - x_Rfi_c)   'center of Rfi on FS xy-cordinate

      ' angle :
            Phi_0 = q_tmp(24)          ' start angle of inner arc    on FS xy-cordinate
            Phi_1 = q_tmp(21)          ' end angle of inner arc      on FS xy-cordinate

            Phi_00 = q_tmp(34)         ' start angle of outer arc    on OS Plate xy-cordinate
            Phi_2 = q_tmp(35)          ' end angle of outer arc      on OS Plate xy-cordinate

        Call Get_Gravity_Center_FS_OS_tg9(Phi_0, Phi_1, Phi_00, Phi_2, the_1)
            Sg_f(9, I) = Area_tmp
            xg_f(9, I) = xg_a_tmp
            yg_f(9, I) = yg_a_tmp
          Call change_Wrap_data_to_curve_xw(0, x_out, y_out)
          Call change_Wrap_data_to_curve_xw(1, x_in, y_in)

      '--Dara Paste -
        Curve_name = "tg9_":    ' DataSheetName = "DataSheet_2"
        Call Paste_curve_data_Num(24, Curve_name & "in", x_in(), y_in(), DataSheetName)
        Call Paste_curve_data_Num(25, Curve_name & "out", x_out(), y_out(), DataSheetName)
'   Stop

        Debug.Print "Sg_f(9, i) ="; Format(Area_tmp, "####.####"); Tab(2); _
                              "xg="; Format(xg_f(9, I), "####.####"); "  "; _
                              "yg="; Format(yg_f(9, I), "####.####")

      '-- Routin to draw OS plate outer circle
      '    Call Get_Gravity_Center_FS_OS_Plate(Phi_0, Phi_1, Phi_00, Phi_2, the_1)


            Debug.Print "      time = " & I & Format(Time - srt_time_1, "  HH:mm:ss")
'            Debug.Print "　Index_i = " & Index_i & Format(Time - srt_time_1, "  HH:mm:ss")
'   GoTo Label_Gravity_center_end


   '■■Label point
Label_Gravity_center_end:


'Next I

            Debug.Print "■ 計算 END time= "; Time
            Debug.Print "　処理時間 = " & Format(Time - srt_time_1, "HH:mm:ss")


End Sub


'======================================================== 【      】
'  関数：Get_Gravity_Center_FS_wrap()
'
'========================================================

Public Sub Get_Gravity_Center_FS_wrap(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                     ByVal Phi_00 As Double, ByVal Phi_2 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'-- FS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'  1) FS outer line
'------------------------------------------------

    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura (7) fo
      xfo(I) = Fn_xfo(phi1_v(I))  '+ dx
      yfo(I) = Fn_yfo(phi1_v(I))  '+ dy

'      xfo(i) = -a * phi1_v(i) ^ k * Cos(phi1_v(i) - qq) + g1 * Cos(phi1_v(i) - qq - Atn(k / phi1_v(i))) '+ dx
'      yfo(i) = -a * phi1_v(i) ^ k * Sin(phi1_v(i) - qq) + g1 * Sin(phi1_v(i) - qq - Atn(k / phi1_v(i))) '+ dy

    Next I

'  2) FS inner line
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I

    ' Formura (8) fi
         xfi(I) = Fn_xfi(phi1_v(I))  '+ dx
         yfi(I) = Fn_yfi(phi1_v(I))  '+ dy

   '   xfi(i) = a * phi1_v(i) ^ k * Cos(phi1_v(i) - qq) + g1 * Cos(phi1_v(i) - qq - Atn(k / phi1_v(i))) '+ dx
   '   yfi(i) = a * phi1_v(i) ^ k * Sin(phi1_v(i) - qq) + g1 * Sin(phi1_v(i) - qq - Atn(k / phi1_v(i))) '+ dy

    Next I

'  3) FS wrap Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp
        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area
        If Area_A <= 0 Then
          xg_a_tmp = 0
          yg_a_tmp = 0
        Else
          xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
          yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
          Area_tmp = Area_A
        End If

' check for curve
        For I = 0 To div_n
              x_out(I) = xfo(I) + dx       ' offset 加える
              y_out(I) = yfo(I) + dy
              x_in(I) = xfi(I) + dx       ' offset 加える
              y_in(I) = yfi(I) + dy
        Next I


End Sub


'======================================================== 【      】
'  関数：Get_Gravity_Center_OS_wrap()
'
'========================================================

Public Sub Get_Gravity_Center_OS_wrap(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                     ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　OS Wrap  ：配列設定
'------------------
'div_n = UBound(the_c)
'div_n = UBound(x_out)

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

    ReDim x_out(div_n):     ReDim y_out(div_n):      ' 表示確認用
    ReDim x_in(div_n):      ReDim y_in(div_n):

'------------------
'-- OS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

    If (Phi_2 = Phi_00) And (Phi_1 = Phi_0) Then
        For I = 0 To div_n
              x_out(I) = 0     ' FS-xy
              y_out(I) = 0     ' FS-xy
              x_in(I) = 0      ' FS-xy
              y_in(I) = 0      ' FS-xy
        Next I

        GoTo Gotoend
    End If


'  1) OS outer line
'------------------------------------------------

    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura (5) mo
         xmo(I) = Fn_xmo(phi1_v(I))  '+ dx
         ymo(I) = Fn_ymo(phi1_v(I))  '+ dy

         xmo(I) = xmo(I) + Ro * Cos(the_1)
         ymo(I) = ymo(I) + Ro * Sin(the_1)

    Next I

'  2) OS inner line
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I

    ' Formura (6) mi
         xmi(I) = Fn_xmi(phi1_v(I))  '+ dx
         ymi(I) = Fn_ymi(phi1_v(I))  '+ dy

         xmi(I) = xmi(I) + Ro * Cos(the_1)
         ymi(I) = ymi(I) + Ro * Sin(the_1)

    Next I

'  3) OS wrap Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                        - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

        del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                        - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- OS wrap Area
        Area_A = (Smo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp

'--- center of area

        If Area_A <= 0 Then
          xg_a_tmp = 0
          yg_a_tmp = 0
        Else
          xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
          yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
          Area_tmp = Area_A
        End If

' check for curve
        For I = 0 To div_n
               x_out(I) = xmo(I) + dx      ' FS-xy
               y_out(I) = ymo(I) + dy      ' FS-xy
               x_in(I) = xmi(I) + dx      ' FS-xy
               y_in(I) = ymi(I) + dy      ' FS-xy
        Next I

'Label
Gotoend:

'div_n = UBound(curve_xw, 2)

End Sub


'======================================================== 【      】
'  関数：Get_Gravity_Center_OS_wrap_all()
'　　　only Wrap part, on the xy of OS-plate
'
'
'========================================================

Public Sub Get_Gravity_Center_OS_wrap_all(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                ByVal Phi_00 As Double, ByVal Phi_2 As Double)  ', ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double
    Dim tmp_N As Long, tmp_div_n As Long

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double       ' 分割面積
    Dim Area_A As Double:                                     ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

    Dim Curve_name As String

'------------------
'　OS Wrap  ：配列設定
'------------------
    tmp_div_n = div_n      ' set division number  分割数の変更

    tmp_N = 3
    div_n = div_n * tmp_N

  '---
    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

    ReDim x_out(div_n):     ReDim y_out(div_n):      ' 表示確認用
    ReDim x_in(div_n):      ReDim y_in(div_n):

'------------------
'-- OS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

'       Phi_0 = 1 ' Phi_c_fi(3, Index_I)       ' in=srt '  Phi_0 < Phi_1 < Phi_2      PI
'       Phi_1 = 2 '  Phi_c_fi(2, Index_I)       ' in=end
'       Phi_00 = 2 '  Phi_c_fi(2, Index_I)       ' out=srt
'       Phi_2 = 4 '  Phi_c_fi(1, Index_I)       ' out=end


'  1) OS outer line
'------------------------------------------------

    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I

    ' Formura (5) mo
         xmo(I) = Fn_xmo(phi1_v(I))  '+ dx
         ymo(I) = Fn_ymo(phi1_v(I))  '+ dy

    Next I

'  2) OS inner line
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I

    ' Formura (6) mi
         xmi(I) = Fn_xmi(phi1_v(I))  '+ dx
         ymi(I) = Fn_ymi(phi1_v(I))  '+ dy

    Next I

'  3) OS wrap Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                        - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

        del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                        - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- OS wrap Area
        Area_A = (Smo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx  '+ Ro * Cos(the_1)　：旋廻移動分なし
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy  '+ Ro * Sin(the_1)
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I) + dx     ' + Ro * Cos(the_1)     ' FS-xy 基準 ：旋廻移動分なし
               y_out(I) = ymo(I) + dy     ' + Ro * Sin(the_1)   ' FS-xy 基準 ：旋廻移動分なし
               x_in(I) = xmi(I) + dx      ' + Ro * Cos(the_1)      ' FS-xy 基準 ：旋廻移動分なし
               y_in(I) = ymi(I) + dy      ' + Ro * Sin(the_1)    ' FS-xy 基準 ：旋廻移動分なし
         Next I

'--- Restore the number to the original division number　分割数を元に戻す
        div_n = tmp_div_n


End Sub




'======================================================== 【      】
'  関数：Get_Gravity_Center_OS_wrap_all()
'　　　only Wrap part, on the xy of OS-plate
'
'
'========================================================

Public Sub Get_Gravity_Center_OS_wrap_all_div2(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                ByVal Phi_00 As Double, ByVal Phi_2 As Double)  ', ByVal the_1 As Double)

    Dim tmp_out_0 As Double:    Dim tmp_out_1 As Double:
    Dim tmp_in_0 As Double:     Dim tmp_in_1 As Double:

'    Dim tmp_N As Double:
'    Dim tmp_q_mr As Double:  Dim tmp_q_mr2 As Double:
'    Dim tmp_Fc As Double:

    Dim tmp_Area(2) As Double:    Dim Area_all As Double:
'    Dim tmp_Mgx As Double:      Dim tmp_Mgy As Double:      Dim tmp_Mgz As Double

'    Dim tmp_V(9) As Double:    Dim tmp_Mg(9) As Double:
    Dim tmp_Gx(2) As Double:   Dim tmp_Gy(2) As Double:   Dim tmp_Gz(2) As Double:

'    Dim I As Long, J As Long
    Dim the_1 As Double:
    Dim Curve_name As String
'
'    Dim I1 As Long, J1 As Long
'    Dim Imax As Long, Jmax As Long


 '------------

        Index_I = 0
        the_1 = the(Index_I) - qq

'------------
' 1)　'[tp1 OS]  巻き始め　吐出室内部分

        tmp_in_0 = Wrap_Start_angle_min(3)        ' see <Calc_Gravity_Center_wrap>
        tmp_in_1 = Phi_c_fi(6, Index_I) - pi * 4 / 6
        tmp_out_0 = Wrap_Start_angle_min(4)       ' see <>'[tp1 OS]   outer Wrap Wall
        tmp_out_1 = Phi_c_fi(5, Index_I) - pi * 5 / 6

            '  tmp_in_0 = OS_in_srt_0        '
            '  tmp_in_1 = OS_in_end
            '  tmp_out_0 = OS_out_srt_0      ' see <>'[tp1 OS]   outer Wrap Wall
            '  tmp_out_1 = OS_out_end

           Call Get_Gravity_Center_OS_wrap(tmp_in_0, tmp_in_1, tmp_out_0, tmp_out_1, the_1)
              '**in,out線, xg_a_tmp,yg_a_tmp 旋廻含む

              tmp_Gx(1) = xg_a_tmp  ' + Ro * Cos(the_1) 含む
              tmp_Gy(1) = yg_a_tmp  ' + Ro * Sin(the_1)
              tmp_Area(1) = Area_tmp

                '--Dara Paste -
                  Curve_name = "OS wrap_1":    ' DataSheetName = "DataSheet_2"
                    Call Paste_curve_data_Num(28, Curve_name & "in", x_in(), y_in(), DataSheetName)
                    Call Paste_curve_data_Num(29, Curve_name & "out", x_out(), y_out(), DataSheetName)
                 ' Stop


'------------
' 2)　'[tp5-tp2 OS]　tp1以外

        tmp_in_0 = Phi_c_fi(6, Index_I) - pi * 4 / 6    '
        tmp_in_1 = OS_in_end                  ' Phi_c_fi(2, Index_I)
        tmp_out_0 = Phi_c_fi(5, Index_I) - pi * 5 / 6    '
        tmp_out_1 = OS_out_end                ' Phi_c_fi(1, Index_I)

         div_n = div_n * 3

           Call Get_Gravity_Center_OS_wrap(tmp_in_0, tmp_in_1, tmp_out_0, tmp_out_1, the_1)
              '**in,out線, xg_a_tmp,yg_a_tmp 旋廻無

              tmp_Gx(2) = xg_a_tmp  ' + Ro * Cos(the_1)
              tmp_Gy(2) = yg_a_tmp  ' + Ro * Sin(the_1)
              tmp_Area(2) = Area_tmp

                '--Dara Paste -
                  Curve_name = "OS wrap_2":    ' DataSheetName = "DataSheet_2"
                    Call Paste_curve_data_Num(30, Curve_name & "in", x_in(), y_in(), DataSheetName)
                    Call Paste_curve_data_Num(31, Curve_name & "out", x_out(), y_out(), DataSheetName)
                 ' Stop

         div_n = div_n / 3

'------------
' 3) total

              tmp_Area(0) = tmp_Area(1) + tmp_Area(2)

        xg_a_tmp = (tmp_Area(1) * tmp_Gx(1) + tmp_Area(2) * tmp_Gx(2)) / (tmp_Area(0))
        yg_a_tmp = (tmp_Area(1) * tmp_Gy(1) + tmp_Area(2) * tmp_Gy(2)) / (tmp_Area(0))
        Area_tmp = tmp_Area(0)



End Sub


'======================================================== 【      】
'      Get_Gravity_Center_Chamber_D_FS
'
'========================================================

Public Sub Get_Gravity_Center_Chamber_D_FS(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                                                        ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'-- FS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'  1) Discharge chamber outer line      : A side : FS inner wall
'------------------------------------------------

    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura (8) fi
         xfi(I) = Fn_xfi(phi1_v(I))  '+ dx
         yfi(I) = Fn_yfi(phi1_v(I))  '+ dy

   '      xfi(i) = a * phi1_v(i) ^ k * Cos(phi1_v(i) - qq) + g1 * Cos(phi1_v(i) - qq - Atn(k / phi1_v(i))) '+ dx
   '      yfi(i) = a * phi1_v(i) ^ k * Sin(phi1_v(i) - qq) + g1 * Sin(phi1_v(i) - qq - Atn(k / phi1_v(i))) '+ dy

    Next I

'  2) Discharge chamber :    A side : FS Center partition lien
'------------------------------------------------
      '直線：(xfo(0),yfo(0)) (xfo(div_n),yfo(div_n))を結ぶ直線

      ' Formura (6)    (xfo(0),yfo(0))=(xfi(0),yfi(0))
        xfo(0) = xfi(0)       '+ dx
        yfo(0) = yfi(0)       '+ dy
        xfo(div_n) = Fn_xmi(phi1_v(0)) + Ro * Cos(the_1) '+ dx
        yfo(div_n) = Fn_ymi(phi1_v(0)) + Ro * Sin(the_1) '+ dy

    For I = 0 To div_n
        xfo(I) = xfo(0) + (xfo(div_n) - xfo(0)) / div_n * I
        yfo(I) = yfo(0) + (yfo(div_n) - yfo(0)) / div_n * I
    Next I

'      '--Dara Paste -
''        Curve_name = "pp1 FS_":    ' DataSheetName = "DataSheet_2"
'        Call Paste_curve_data_Num(0, "pp1 FS_in", xfi(), yfi(), DataSheetName)
'        Call Paste_curve_data_Num(1, "pp1 FS_out", xfo(), yfo(), DataSheetName)
'   Stop

'  3) FS wrap Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp

        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I) + dx
               y_out(I) = yfo(I) + dy

               x_in(I) = xfi(I) + dx
               y_in(I) = yfi(I) + dy
         Next I

End Sub


'======================================================== 【      】
'  Get_Gravity_Center_Chamber_D_OS
'
'========================================================

Public Sub Get_Gravity_Center_Chamber_D_OS(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                                                        ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　OS Wrap  ：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)


'------------------
'-- OS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

'  1) Discharge chamber outer line      : B side : OS inner line
'------------------------------------------------

    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I

      ' Formura (6) mi
        xmi(I) = Fn_xmi(phi1_v(I))  '+ dx
        ymi(I) = Fn_ymi(phi1_v(I))  '+ dy

'            xmi(i) = -a * phi1_v(i) ^ k * Cos(phi1_v(i) - qq) - g2 * Cos(phi1_v(i) - qq - Atn(k / phi1_v(i))) '+ dx
'            ymi(i) = -a * phi1_v(i) ^ k * Sin(phi1_v(i) - qq) - g2 * Sin(phi1_v(i) - qq - Atn(k / phi1_v(i))) '+ dy

        xmi(I) = xmi(I) + Ro * Cos(the_1)
        ymi(I) = ymi(I) + Ro * Sin(the_1)

    Next I

'  2) Discharge chamber  :    A side : OS Center partition lien
'------------------------------------------------
    '直線：(xmo(0),ymo(0)) (xmo(div_n),ymo(div_n))を結ぶ直線
    ' Formura (8) fi=mi
         xmo(0) = xmi(0)      '+ dx
         ymo(0) = xmi(0)      '+ dy
         xmo(div_n) = Fn_xfi(phi1_v(0))  '+ dx
         ymo(div_n) = Fn_yfi(phi1_v(0))  '+ dy

     For I = 0 To div_n
        xmo(I) = xmi(0) + (xmo(div_n) - xmo(0)) / div_n * I
        ymo(I) = ymi(0) + (ymo(div_n) - ymo(0)) / div_n * I
     Next I

'      '--Dara Paste -
'        Call Paste_curve_data_Num(2, "pg1 OS_in", xmi(), ymi(), DataSheetName)
'        Call Paste_curve_data_Num(3, "pg1 OS_out", xmo(), ymo(), DataSheetName)
'   Stop


'  3) FS wrap Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                        - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

        del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                        - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Smo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp

        Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
        Area_tmp = Area_A

         For I = 0 To div_n
               x_out(I) = xmo(I) + dx
               y_out(I) = ymo(I) + dy

               x_in(I) = xmi(I) + dx
               y_in(I) = ymi(I) + dy
         Next I

End Sub

'======================================================== 【      】
'  関数：Get_Gravity_Center_Chamber_A
'                      vs Get_Gravity_Center_FS_wrap()
'
'========================================================

Public Sub Get_Gravity_Center_Chamber_A(ByVal Phi_1 As Double, ByVal Phi_2 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

    Dim the_1 As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)


'------------------
'-- A Chamber
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Smo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

   the_1 = Phi_1 - Atn(k / Phi_1) - qq

'  1) FS inner line
'------------------------------------------------
    div_phi = (Phi_2 - Phi_1) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_1 + div_phi * I
    ' Formura (8) fi
      xfi(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
      yfi(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

    Next I

'  2) OS outer line
'------------------------------------------------

    div_phi = (Phi_2 - Phi_1) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_1 + div_phi * I
    ' Formura (7) fo
        xmo(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) - g2 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        ymo(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) - g2 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

        xmo(I) = xmo(I) + Ro * Cos(the_1)
        ymo(I) = ymo(I) + Ro * Sin(the_1)

    Next I


'  3) A Chamber Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-mo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xmo(I - 1)) * (yfi(I - 1) - ymo(I - 1)) _
                        - (xfi(I - 1) - xmo(I - 1)) * (yfi(I) - ymo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xmo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + ymo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-fi(i)]

        del_Smo = Abs((xmo(I) - xfi(I)) * (ymo(I - 1) - yfi(I)) _
                        - (xmo(I - 1) - xfi(I)) * (ymo(I) - yfi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xfi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + yfi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Smo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_A = Sgx_mo_tmp + Sgx_fi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_fi_tmp

'--- center of area
        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I) + dx      ' FS-xy
               y_out(I) = ymo(I) + dy      ' FS-xy
               x_in(I) = xfi(I) + dx      ' FS-xy
               y_in(I) = yfi(I) + dy      ' FS-xy
         Next I

End Sub


'======================================================== 【      】
'  関数：Get_Gravity_Center_Chamber_B
'                  vs   Get_Gravity_Center_OS_wrap()
'
'========================================================

Public Sub Get_Gravity_Center_Chamber_B(ByVal Phi_1 As Double, ByVal Phi_2 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' B : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　OS Wrap  ：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)

    Dim the_1 As Double

'------------------
'-- OS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Smi() , del_Sfo()
'------------------------------------------------
    Smi_tmp = 0
    Sfo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

   the_1 = Phi_1 - Atn(k / Phi_1) - qq


'  1) OS inner line
'------------------------------------------------
    div_phi = (Phi_2 - Phi_1) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_1 + div_phi * I
    ' Formura (6) mi
        xmi(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) - g2 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        ymi(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) - g2 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

        xmi(I) = xmi(I) + Ro * Cos(the_1)
        ymi(I) = ymi(I) + Ro * Sin(the_1)

    Next I

'  2) FS outer line
'------------------------------------------------
    div_phi = (Phi_2 - Phi_1) / div_n    ' Divied angle　分割の角度幅

     For I = 0 To div_n
        phi1_v(I) = Phi_1 + div_phi * I
     ' Formura (7) fo
        xfo(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        yfo(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

    Next I


'  3) B Chamber Area Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-fo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xfo(I - 1)) * (ymi(I - 1) - yfo(I - 1)) _
                        - (xmi(I - 1) - xfo(I - 1)) * (ymi(I) - yfo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xfo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + yfo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-mi(i)]

        del_Sfo = Abs((xfo(I) - xmi(I)) * (yfo(I - 1) - ymi(I)) _
                        - (xfo(I - 1) - xmi(I)) * (yfo(I) - ymi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xmi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + ymi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_A = Sgx_fo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_fo_tmp + Sgy_mi_tmp

'--- center of area
        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I) + dx      ' FS-xy
               y_out(I) = yfo(I) + dy      ' FS-xy
               x_in(I) = xmi(I) + dx      ' FS-xy
               y_in(I) = ymi(I) + dy      ' FS-xy
         Next I

End Sub


'======================================================== 【      】
'  関数：Get_Gravity_Center_Chamber_Suction
'
'
'========================================================

Public Sub Get_Gravity_Center_Chamber_Suction(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                             ByVal Phi_00 As Double, ByVal Phi_2 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double
    Dim the_1 As Double

    Dim Sfi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

    Dim Smi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_B As Double:

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

' B : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_B As Double:         Dim Sgy_Area_B As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'------------------
'　OS Wrap  ：配列設定
    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)

'------------------

    Sfi_tmp = 0
    Smo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

    Smi_tmp = 0
    Sfo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

   the_1 = Phi_1 - Atn(k / Phi_1) - qq


'===============================================
'------------------------------------------------
'  1) A-side :FS inner line
'------------------------------------------------
   If Phi_00 <> Phi_2 Then

       div_phi = (Phi_00 - Phi_2) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_2 + div_phi * I
      ' Formura (8) fi
        xfi(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        yfi(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

      ' Formura (16) mo
        xmo(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) - g2 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        ymo(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) - g2 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

        xmo(I) = xmo(I) + Ro * Cos(the_1)
        ymo(I) = ymo(I) + Ro * Sin(the_1)

    Next I


   '  3) A-side : A-Chamber Area  and Gravity center
   '------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-mo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xmo(I - 1)) * (yfi(I - 1) - ymo(I - 1)) _
                        - (xfi(I - 1) - xmo(I - 1)) * (yfi(I) - ymo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xmo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + ymo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-fi(i)]

        del_Smo = Abs((xmo(I) - xfi(I)) * (ymo(I - 1) - yfi(I)) _
                        - (xmo(I - 1) - xfi(I)) * (ymo(I) - yfi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xfi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + yfi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

   '---- FS wrap Area
        Area_A = (Smo_tmp + Sfi_tmp)

   '---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_A = Sgx_mo_tmp + Sgx_fi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_fi_tmp

   '--- center of area
        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
        Area_tmp = Area_A

  Else
        xg_a_tmp = 0
        yg_a_tmp = 0
        Area_tmp = 0

  End If

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I) + dx      ' FS-xy
               y_out(I) = ymo(I) + dy      ' FS-xy
               x_in(I) = xfi(I) + dx      ' FS-xy
               y_in(I) = yfi(I) + dy      ' FS-xy
         Next I


'===============================================
'  1)B- side : OS inner & FS outer line
'------------------------------------------------
    div_phi = (Phi_0 - Phi_1) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_1 + div_phi * I
    ' Formura (6) mi
        xmi(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) - g2 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        ymi(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) - g2 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

        xmi(I) = xmi(I) + Ro * Cos(the_1)
        ymi(I) = ymi(I) + Ro * Sin(the_1)

   ' Formura (7) fo
        xfo(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        yfo(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

    Next I



'  3)B- side: B-Chamber Area Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-fo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xfo(I - 1)) * (ymi(I - 1) - yfo(I - 1)) _
                        - (xmi(I - 1) - xfo(I - 1)) * (ymi(I) - yfo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xfo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + yfo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-mi(i)]

        del_Sfo = Abs((xfo(I) - xmi(I)) * (yfo(I - 1) - ymi(I)) _
                        - (xfo(I - 1) - xmi(I)) * (yfo(I) - ymi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xmi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + ymi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_B = (Sfo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_B = Sgx_fo_tmp + Sgx_mi_tmp
        Sgy_Area_B = Sgy_fo_tmp + Sgy_mi_tmp

'--- center of area
        xg_a_tmp = (Sgy_Area_B) / (Area_B) + dx
        yg_a_tmp = (Sgx_Area_B) / (Area_B) + dy
        Area_tmp = Area_B

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I) + dx      ' FS-xy
               y_out(I) = yfo(I) + dy      ' FS-xy
               x_in(I) = xmi(I) + dx      ' FS-xy
               y_in(I) = ymi(I) + dy      ' FS-xy
         Next I


'===============================================
'  1) A+B side -> A
'------------------------------------------------

''---- FS wrap Area
'        Area_A = Area_A + Area_B
'
''---  area geometrical moment of area Sgx, Sgy 総和
'        Sgx_Area_A = Sgx_Area_A + Sgx_Area_B
'        Sgy_Area_A = Sgy_Area_A + Sgy_Area_B
'
''--- center of area
'        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
'        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
'        Area_tmp = Area_A




End Sub

'======================================================== 【      】
'  関数：Get_Gravity_Center_Chamber_Suction_1B
'
'
'========================================================

Public Sub Get_Gravity_Center_Chamber_Suction_1B(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                             ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)
    Dim I As Long, J As Long
    Dim div_phi As Double
'    Dim the_1 As Double

    Dim Sfi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

    Dim Smi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_B As Double:

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

' B : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_B As Double:         Dim Sgy_Area_B As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'------------------
'　OS Wrap  ：配列設定
    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)

'------------------

    Sfi_tmp = 0
    Smo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

    Smi_tmp = 0
    Sfo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0


'------------------------------------------------
'  1)  Suction area-1 *B chamber side : OS inner
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura (6) mi
        xmi(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) - g2 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        ymi(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) - g2 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy
            xmi(I) = xmi(I) + Ro * Cos(the_1)
            ymi(I) = ymi(I) + Ro * Sin(the_1)
    Next I

'  2)  Suction area-1 *B chamber side : FS outer line
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
   ' Formura (7) fo
        xfo(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        yfo(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

    Next I

'  3) Suction area-1 *B Chamber : area and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-fo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xfo(I - 1)) * (ymi(I - 1) - yfo(I - 1)) _
                        - (xmi(I - 1) - xfo(I - 1)) * (ymi(I) - yfo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xfo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + yfo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-mi(i)]

        del_Sfo = Abs((xfo(I) - xmi(I)) * (yfo(I - 1) - ymi(I)) _
                        - (xfo(I - 1) - xmi(I)) * (yfo(I) - ymi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xmi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + ymi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_A = Sgx_fo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_fo_tmp + Sgy_mi_tmp

'--- center of area
        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I) + dx      ' FS-xy
               y_out(I) = yfo(I) + dy      ' FS-xy
               x_in(I) = xmi(I) + dx      ' FS-xy
               y_in(I) = ymi(I) + dy      ' FS-xy
         Next I

End Sub

'======================================================== 【      】
'  関数：Get_Gravity_Center_Chamber_Suction_1A
'
'
'========================================================

Public Sub Get_Gravity_Center_Chamber_Suction_1A(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                                             ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)
    Dim I As Long, J As Long
    Dim div_phi As Double
'    Dim the_1 As Double

    Dim Sfi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

    Dim Smi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_B As Double:

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

' B : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_B As Double:         Dim Sgy_Area_B As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'------------------
'　OS Wrap  ：配列設定
    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)

'------------------

    Sfi_tmp = 0
    Smo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

    Smi_tmp = 0
    Sfo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0


'------------------------------------------------
'  1)  Suction area-1 *A chamber side : FS inner
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura (8) fi
      xfi(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
      yfi(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

    Next I

'  2)  Suction area-1 *A chamber side : OS outer line
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura (7) fo
        xmo(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) - g2 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
        ymo(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) - g2 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

        xmo(I) = xmo(I) + Ro * Cos(the_1)
        ymo(I) = ymo(I) + Ro * Sin(the_1)

    Next I

'  3) Suction area-1 *B Chamber : area and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-fo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xfo(I - 1)) * (ymi(I - 1) - yfo(I - 1)) _
                        - (xmi(I - 1) - xfo(I - 1)) * (ymi(I) - yfo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xfo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + yfo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-mi(i)]

        del_Sfo = Abs((xfo(I) - xmi(I)) * (yfo(I - 1) - ymi(I)) _
                        - (xfo(I - 1) - xmi(I)) * (yfo(I) - ymi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xmi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + ymi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_A = Sgx_fo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_fo_tmp + Sgy_mi_tmp

'--- center of area
'        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
'        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
'        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I) + dx      ' FS-xy
               y_out(I) = ymo(I) + dy      ' FS-xy
               x_in(I) = xfi(I) + dx      ' FS-xy
               y_in(I) = yfi(I) + dy      ' FS-xy
         Next I


End Sub

'======================================================== 【      】
'  関数：Get_Gravity_Center_Chamber_Suction_2
'
'
'========================================================

Public Sub Get_Gravity_Center_Chamber_Suction_2(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                          ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)
    Dim I As Long, J As Long
    Dim div_phi As Double
'    Dim the_1 As Double

    Dim Sfi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

    Dim Smi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_B As Double:

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

' B : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_B As Double:         Dim Sgy_Area_B As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'------------------
'　OS Wrap  ：配列設定
    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)

'------------------

    Sfi_tmp = 0
    Smo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

    Smi_tmp = 0
    Sfo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'------------------------------------------------
'  1) Suction area-2 : FS inner
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura (8) fi
      xfi(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
      yfi(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

    Next I

'  2) Suction area-2 : FS outer line
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura (7) fo
      xfo(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
      yfo(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

    Next I

'  3) Suction area-2 : area and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp

        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I) + dx      ' FS-xy
               y_out(I) = yfo(I) + dy      ' FS-xy
               x_in(I) = xfi(I) + dx      ' FS-xy
               y_in(I) = yfi(I) + dy      ' FS-xy
         Next I

'===============================================
'  1) A+B side -> A
'------------------------------------------------

''---- FS wrap Area
'        Area_A = Area_A + Area_B
'
''---  area geometrical moment of area Sgx, Sgy 総和
'        Sgx_Area_A = Sgx_Area_A + Sgx_Area_B
'        Sgy_Area_A = Sgy_Area_A + Sgy_Area_B
'
''--- center of area
'        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
'        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
'        Area_tmp = Area_A

End Sub

'======================================================== 【      】
'  関数：Get_Gravity_Center_Chamber_Suction_3
'
'
'========================================================
Public Sub Get_Gravity_Center_Chamber_Suction_3(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim dmmy_a As Double:    Dim dmmy_b As Double:    Dim dmmy_c As Double:
    Dim dmmy_x As Double:    Dim dmmy_y As Double:
    Dim dmmy_q As Double:    Dim dmmy_R As Double:

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double


'------------------
'-- FS Wrap
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'------------------
'　FS Wrap　　：配列設定
'------------------
    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)

'------------------------------------------------
'  1) outer arc: radius = 57.14  Center (xc,yc)=(4.47,3.32)  inlet FS inner line :
'------------------------------------------------
   I = 0
      phi1_v(0) = Phi_00
    ' Formura (8) fi   on FS Wrap-xy-cordinate
      xfi(0) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
      yfi(0) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

     ' r_Rfi_c = Sqr((yfi(0) - y_Rfi_c) ^ 2 + (xfi(0) - x_Rfi_c) ^ 2)

  ' start angle : Function Atan2(y As Double, x As Double)  on FS Wrap-xy-cordinate
      Phi_00 = Atan2((yfi(0) + dy - y_Rfi_c), (xfi(0) + dx - x_Rfi_c))

  '-----
  ' start Point of inner arc
      x_tmp(27) = xfi(0) + dx     'on FS xy-cordinate
      y_tmp(27) = yfi(0) + dy     'on FS xy-cordinate
      q_tmp(27) = Phi_00          'on FS Wrap-xy-cordinate

  ' arc end point : on FS Wrap-xy-cordinate
      If Phi_2 = pi / 2 Then
        xfi(div_n) = 0
        yfi(div_n) = Sqr(r_Rfi_c ^ 2 - ((x_Rfi_c) ^ 2 + (y_Rfi_c) ^ 2)) + y_Rfi_c

      ElseIf Phi_2 = -pi / 2 Then
        xfi(div_n) = 0
        yfi(div_n) = -Sqr(r_Rfi_c ^ 2 - ((x_Rfi_c) ^ 2 + (y_Rfi_c) ^ 2)) + y_Rfi_c

      Else
        dmmy_q = Tan(Phi_2)
        dmmy_a = 1 + dmmy_q ^ 2
        dmmy_b = -2 * (x_Rfi_c + dmmy_q * y_Rfi_c)
        dmmy_c = ((x_Rfi_c) ^ 2 + (y_Rfi_c) ^ 2) - r_Rfi_c ^ 2

        xfi(div_n) = (-dmmy_b + Sqr(dmmy_b ^ 2 - 4 * dmmy_a * dmmy_c)) / (2 * dmmy_a)
        yfi(div_n) = dmmy_q * xfi(div_n)

      End If

  ' end angle : Phi_2  on FS Wrap-xy-cordinate

       Phi_2 = Atan2((yfi(div_n) + dy - y_Rfi_c), (xfi(div_n) + dx - x_Rfi_c))
          If Phi_2 < Phi_00 Then
              Phi_2 = Phi_2 + pi
          End If

  '-----
  ' end Point of outer arc
      x_tmp(21) = xfi(div_n) + dx     'on FS xy-cordinate
      y_tmp(21) = yfi(div_n) + dy     'on FS xy-cordinate
      q_tmp(21) = Phi_2               'on FS Wrap-xy-cordinate

  '-----
      div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

        For I = 0 To div_n
            phi1_v(I) = Phi_00 + div_phi * I
        ' Formura arc on fi side  on FS Wrap-xy-cordinate
          xfi(I) = r_Rfi_c * Cos(phi1_v(I)) + x_Rfi_c - dx
          yfi(I) = r_Rfi_c * Sin(phi1_v(I)) + y_Rfi_c - dy

        Next I


'------------------------------------------------
'  2) inner arc : radius = 44.1  Center (xc,yc)=(6.02,1.2)   inlet FS outer line
'------------------------------------------------
   I = 0
        phi1_v(I) = Phi_0
    ' Formura (7) fo
      xfo(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dx
      yfo(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) '+ dy

   ' start angle : Function Atan2(y As Double, x As Double)  on FS Wrap-xy-cordinate
      Phi_0 = Atan2((yfo(0) + dy - y_Rfo_c), (xfo(0) + dx - x_Rfo_c)) ' + pi

  '-----
  ' start Point of inner arc
      x_tmp(25) = xfo(0) + dx     'on FS xy-cordinate
      y_tmp(25) = yfo(0) + dy     'on FS xy-cordinate
      q_tmp(25) = Phi_0           'on FS Wrap-xy-cordinate

   ' arc end point : on FS Wrap-xy-cordinate
   '     cross point between arc and line
   '
      If Phi_1 = pi / 2 Then
        xfo(div_n) = 0
        yfo(div_n) = Sqr(r_Rfo_c ^ 2 - ((x_Rfo_c) ^ 2 + (y_Rfo_c) ^ 2)) + y_Rfo_c

      ElseIf Phi_1 = -pi / 2 Then
        xfo(div_n) = 0
        yfo(div_n) = -Sqr(r_Rfo_c ^ 2 - ((x_Rfo_c) ^ 2 + (y_Rfo_c) ^ 2)) + y_Rfo_c

      Else
        dmmy_q = Tan(Phi_1)
        dmmy_a = 1 + dmmy_q ^ 2
        dmmy_b = -2 * (x_Rfo_c + dmmy_q * y_Rfo_c)
        dmmy_c = ((x_Rfo_c) ^ 2 + (y_Rfo_c) ^ 2) - r_Rfo_c ^ 2

        xfo(div_n) = (-dmmy_b + Sqr(dmmy_b ^ 2 - 4 * dmmy_a * dmmy_c)) / (2 * dmmy_a)
        yfo(div_n) = dmmy_q * xfo(div_n)

      End If

  ' end angle Phi_1 : Function Atan2(y As Double, x As Double)  on FS Wrap-xy-cordinate
      Phi_1 = Atan2((yfo(div_n) + dy - y_Rfo_c), (xfo(div_n) + dx - x_Rfo_c))
          If Phi_1 < Phi_0 Then
              Phi_1 = Phi_1 + pi
          End If

  '-----
  ' end Point of inner arc
      x_tmp(22) = xfi(div_n) + dx     'on FS xy-cordinate
      y_tmp(22) = yfi(div_n) + dy     'on FS xy-cordinate
      q_tmp(22) = Phi_1               'on FS Wrap-xy-cordinate

  '-----
      div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

      For I = 0 To div_n
           phi1_v(I) = Phi_0 + div_phi * I
       ' Formura arc on fo side  on FS Wrap-xy-cordinate
         xfo(I) = r_Rfo_c * Cos(phi1_v(I)) + x_Rfo_c - dx
         yfo(I) = r_Rfo_c * Sin(phi1_v(I)) + y_Rfo_c - dy

      Next I


'  3) FS wrap Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp
        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area
        xg_a_tmp = (Sgy_Area_A) / (Area_A) + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) + dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I) + dx       ' FS-xy
               y_out(I) = yfo(I) + dy       ' FS-xy
               x_in(I) = xfi(I) + dx        ' FS-xy
               y_in(I) = yfi(I) + dy       ' FS-xy
         Next I

End Sub


'======================================================== 【      】
'     Get_Gravity_Center_Oilgroove_PW_1(Phi_0, Phi_1, Phi_00, Phi_2)
'
'
'========================================================
Public Sub Get_Gravity_Center_Oilgroove_PW_1(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double
'    Dim the_1 As Double:

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double


'------------------
'-- FS Wrap
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'------------------
'　FS Wrap　　：配列設定
'------------------
    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)

'------------------------------------------------
'  1)  oil groove outer arc1 : radius = 53.72  Center (xc,yc)=(0.89,2.4)
'------------------------------------------------
'   ' Function Atan2(y As Double, x As Double)
'      Phi_00 = Atan2((yfi(0) - y_Rfi_c), (xfi(0) - x_Rfi_c))

      div_phi = (Phi_2 - Phi_00) / div_n    'outer arc Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xfo(I) = r1_oilgroove * Cos(phi1_v(I)) + x1_oilgroove_c
      yfo(I) = r1_oilgroove * Sin(phi1_v(I)) + y1_oilgroove_c

    Next I

  ' start Point of outer arc
      x_tmp(11) = xfo(0) + dx     'on FS xy-cordinate
      y_tmp(11) = yfo(0) + dy     'on FS xy-cordinate
      q_tmp(11) = Phi_00

  ' end Point of outer arc
      x_tmp(12) = xfo(div_n) + dx     'on FS xy-cordinate
      y_tmp(12) = yfo(div_n) + dy     'on FS xy-cordinate
      q_tmp(12) = Phi_2

'------------------------------------------------
'  2) oil groove inner line : radius = 52.22  Center (xc,yc)=(0.89,2.4)
'------------------------------------------------

      div_phi = (Phi_1 - Phi_0) / div_n    ' inner arc Divied angle　分割の角度幅

   For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura arc innner side
      xfi(I) = (r1_oilgroove - t1_oilgroove) * Cos(phi1_v(I)) + x1_oilgroove_c
      yfi(I) = (r1_oilgroove - t1_oilgroove) * Sin(phi1_v(I)) + y1_oilgroove_c

   Next I

  ' start Point of inner arc
      x_tmp(1) = xfi(0) + dx     'on FS xy-cordinate
      y_tmp(1) = yfi(0) + dy     'on FS xy-cordinate
      q_tmp(1) = Phi_0

  ' end Point of inner arc
      x_tmp(2) = xfi(div_n) + dx     'on FS xy-cordinate
      y_tmp(2) = yfi(div_n) + dy     'on FS xy-cordinate
      q_tmp(2) = Phi_1

'------------------------------------------------
'  3)  Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- total Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和
        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp
        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area
        xg_a_tmp = (Sgy_Area_A) / (Area_A)
        yg_a_tmp = (Sgx_Area_A) / (Area_A)
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I)    '+ dx       ' FS-xy
               y_out(I) = yfo(I)    ' + dy       ' FS-xy
               x_in(I) = xfi(I)    ' + dx        ' FS-xy
               y_in(I) = yfi(I)    ' + dy       ' FS-xy
         Next I

End Sub



'======================================================== 【      】
'     Get_Gravity_Center_Oilgroove_PW_2(Phi_0, Phi_1, Phi_00, Phi_2)
'
'
'========================================================
Public Sub Get_Gravity_Center_Oilgroove_PW_2(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　Curve  ：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'------------------
'-- Curve
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

'------------------------------------------------
'  1)  oil groove outer arc2 : radius = 56.35  Center (xc,yc)=(-1.05,-0.25)
'------------------------------------------------
'   ' Function Atan2(y As Double, x As Double)
'      Phi_00 = Atan2((yfi(0) - y_Rfi_c), (xfi(0) - x_Rfi_c))

      div_phi = (Phi_2 - Phi_00) / div_n    '  outer arc Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xmo(I) = r2_oilgroove * Cos(phi1_v(I)) + x2_oilgroove_c '- dx
      ymo(I) = r2_oilgroove * Sin(phi1_v(I)) + y2_oilgroove_c '- dy

    Next I

  ' start Point of outer arc
      x_tmp(13) = xmo(0) + dx     'on FS xy-cordinate
      y_tmp(13) = ymo(0) + dy     'on FS xy-cordinate
      q_tmp(13) = Phi_00

  ' end Point of outer arc
      x_tmp(14) = xmo(div_n) + dx     'on FS xy-cordinate
      y_tmp(14) = ymo(div_n) + dy     'on FS xy-cordinate
      q_tmp(14) = Phi_2

'------------------------------------------------
'  2) oil groove inner arc2 : radius = 56.35-1.5  Center (xc,yc)=(-1.05,-0.25)
'------------------------------------------------

      div_phi = (Phi_1 - Phi_0) / div_n    ' inner arc Divied angle　分割の角度幅

   For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura arc innner side
      xmi(I) = (r2_oilgroove - t2_oilgroove) * Cos(phi1_v(I)) + x2_oilgroove_c '- dx
      ymi(I) = (r2_oilgroove - t2_oilgroove) * Sin(phi1_v(I)) + y2_oilgroove_c '- dy

   Next I

  ' start Point of inner arc
      x_tmp(5) = xmi(0) + dx     'on FS xy-cordinate
      y_tmp(5) = ymi(0) + dy     'on FS xy-cordinate
      q_tmp(5) = Phi_0

  ' end Point of inner arc
      x_tmp(9) = xmi(div_n) + dx     'on FS xy-cordinate
      y_tmp(9) = ymi(div_n) + dy     'on FS xy-cordinate
      q_tmp(9) = Phi_1

'  3) oil groove Area and Gravity center
'------------------------------------------------
     For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                        - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

        del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                        - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- total Area
        Area_A = (Smo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A) '+ dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) '+ dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I)    '+ dx       ' FS-xy
               y_out(I) = ymo(I)    ' + dy       ' FS-xy
               x_in(I) = xmi(I)    ' + dx        ' FS-xy
               y_in(I) = ymi(I)    ' + dy       ' FS-xy
         Next I


End Sub



'======================================================== 【      】
'   Get_Gravity_Center_Oilgroove_OS_1
'   [tg10-1]   right area1 of between oilgroove and OS Plate
'========================================================

Public Sub Get_Gravity_Center_Oilgroove_OS_tg10_1(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                             ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'------------------------------------------------
'  1) outer arc1 : OS-Plate  (use format of FS outer line)
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xfo(I) = OS_dia / 2 * Cos(phi1_v(I)) + Ro * Cos(the_1)
      yfo(I) = OS_dia / 2 * Sin(phi1_v(I)) + Ro * Sin(the_1)

    Next I

  ' start Point of outer arc
      x_tmp(31) = xfo(0)      'on FS xy-cordinate
      y_tmp(31) = yfo(0)      'on FS xy-cordinate
      q_tmp(31) = Phi_00

  ' end Point of outer arc
      x_tmp(32) = xfo(div_n)      'on FS xy-cordinate
      y_tmp(32) = yfo(div_n)      'on FS xy-cordinate
      q_tmp(32) = Phi_2

'------------------------------------------------
'  2) inner arc1 : outer oilgroove   (use format of FS inner line)
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura (8) fi
      xfi(I) = r1_oilgroove * Cos(phi1_v(I)) + x1_oilgroove_c
      yfi(I) = r1_oilgroove * Sin(phi1_v(I)) + y1_oilgroove_c
    Next I


'------------------------------------------------
'  3) Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp

        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A)
        yg_a_tmp = (Sgx_Area_A) / (Area_A)
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I)    '+ dx       ' FS-xy
               y_out(I) = yfo(I)    ' + dy       ' FS-xy
               x_in(I) = xfi(I)    ' + dx        ' FS-xy
               y_in(I) = yfi(I)    ' + dy       ' FS-xy
         Next I


End Sub


'======================================================== 【      】
'    Get_Gravity_Center_Oilgroove_OS_2
'     [tg10-2]  left area2 of between oilgroove and OS Plate

'========================================================
Public Sub Get_Gravity_Center_Oilgroove_OS_tg10_2(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)
    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　OS Wrap  ：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'------------------
'-- OS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

'------------------------------------------------
'  1) OS-Plate outer arc
'------------------------------------------------
'   ' Function Atan2(y As Double, x As Double)
'      Phi_00 = Atan2((yfi(0) - y_Rfi_c), (xfi(0) - x_Rfi_c))

      div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side      'on FS xy-cordinate
      xmo(I) = OS_dia / 2 * Cos(phi1_v(I)) + Ro * Cos(the_1)
      ymo(I) = OS_dia / 2 * Sin(phi1_v(I)) + Ro * Sin(the_1)

    Next I

  ' start Point of outer arc
      x_tmp(33) = xmo(0)      'on FS xy-cordinate
      y_tmp(33) = ymo(0)      'on FS xy-cordinate
      q_tmp(33) = Phi_00

  ' end Point of outer arc
      x_tmp(34) = xmo(div_n)      'on FS xy-cordinate
      y_tmp(34) = ymo(div_n)      'on FS xy-cordinate
      q_tmp(34) = Phi_2


'------------------------------------------------
'  2) oilgroove outer arc1 (FS) inner line
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura arc outer side     'on FS xy-cordinate
      xmi(I) = r2_oilgroove * Cos(phi1_v(I)) + x2_oilgroove_c '- dx
      ymi(I) = r2_oilgroove * Sin(phi1_v(I)) + y2_oilgroove_c '- dy

    Next I


'------------------------------------------------
'  3) Area and Gravity center
'------------------------------------------------
     For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                        - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

        del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                        - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- OS wrap Area
        Area_A = (Smo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A) '+ dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) '+ dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I)  '+ dx      ' FS-xy
               y_out(I) = ymo(I)  '+ dy      ' FS-xy
               x_in(I) = xmi(I)   '+ dx      ' FS-xy
               y_in(I) = ymi(I)   '+ dy      ' FS-xy
         Next I

End Sub

'======================================================== 【      】
'    Get_Gravity_Center_Oilgroove_FS_tg7_1
'     [tg7-1] : [area of Between oil groove inner arc1 and FS_in ]
'             right area1 of between oilgroove and FS
'
'========================================================
Public Sub Get_Gravity_Center_Oilgroove_FS_tg7_1(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0



  If Phi_0 <> Phi_1 Then

    '------------------------------------------------
    '  1) outer arc1 : inner oli groove arc1  (use format of FS outer line)
    '------------------------------------------------

        div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

          For I = 0 To div_n
              phi1_v(I) = Phi_00 + div_phi * I
          ' Formura arc outer side
            xfo(I) = (r1_oilgroove - t1_oilgroove) * Cos(phi1_v(I)) + x1_oilgroove_c
            yfo(I) = (r1_oilgroove - t1_oilgroove) * Sin(phi1_v(I)) + y1_oilgroove_c

          Next I

    '------------------------------------------------
    '  2) innner arc : FS inner curve
    '------------------------------------------------
        div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

        For I = 0 To div_n
            phi1_v(I) = Phi_0 + div_phi * I
        ' Formura (8) fi
            Call Wp_xyfi(phi1_v(I))
                xfi(I) = Wp_xfi          '
                yfi(I) = Wp_yfi
    '      xfi(i) = a * phi1_v(i) ^ k * Cos(phi1_v(i) - qq) + g1 * Cos(phi1_v(i) - qq - Atn(k / phi1_v(i))) + dx
    '      yfi(i) = a * phi1_v(i) ^ k * Sin(phi1_v(i) - qq) + g1 * Sin(phi1_v(i) - qq - Atn(k / phi1_v(i))) + dy

        Next I

      ' start Point of outer arc
          x_tmp(3) = xfi(0) '+ dx     'on FS xy-cordinate
          y_tmp(3) = yfi(0) '+ dy     'on FS xy-cordinate
          q_tmp(3) = Phi_0

      ' end Point of outer arc
          x_tmp(4) = xfi(div_n) '+ dx     'on FS xy-cordinate
          y_tmp(4) = yfi(div_n) '+ dy     'on FS xy-cordinate
          q_tmp(4) = Phi_1

    '------------------------------------------------
    '  3) Area  and Gravity center
    '------------------------------------------------
        For I = 1 To div_n

        ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

            del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                            - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
            Sfi_tmp = Sfi_tmp + del_Sfi

            xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
            yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

            Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
            Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

        ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

            del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                            - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
            Sfo_tmp = Sfo_tmp + del_Sfo

            xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
            yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

            Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
            Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

        Next I

    '---- FS wrap Area
            Area_A = (Sfo_tmp + Sfi_tmp)

    '---  area geometrical moment of area Sgx, Sgy 総和

            Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp

            Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

    '--- center of area

            xg_a_tmp = (Sgy_Area_A) / (Area_A) ' + dx
            yg_a_tmp = (Sgx_Area_A) / (Area_A) ' + dy
            Area_tmp = Area_A

  Else
        For I = 0 To div_n
          xfo(I) = 0
          yfo(I) = 0
          xfi(I) = Wp_xfi          '
          yfi(I) = Wp_yfi

        Next I
          xg_a_tmp = 0
          yg_a_tmp = 0
          Area_tmp = 0
  End If


''--- ' for debug of drowing line
'    For i = 0 To div_n
'      xfo2(i) = xfo(i)
'      yfo2(i) = yfo(i)
'      xfi2(i) = xfi(i)
'      yfi2(i) = yfi(i)
'    Next i

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I)    '+ dx       ' FS-xy
               y_out(I) = yfo(I)    ' + dy       ' FS-xy
               x_in(I) = xfi(I)    ' + dx        ' FS-xy
               y_in(I) = yfi(I)    ' + dy       ' FS-xy
         Next I


End Sub


'======================================================== 【      】
'    Get_Gravity_Center_Oilgroove_FS_tg7_2
'     [tg7-2] : [area of Between oil groove inner arc1 and FS_in ]
'     left area2 of between oilgroove and FS
'
'========================================================

Public Sub Get_Gravity_Center_Oilgroove_FS_tg7_2(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)
    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　OS Wrap  ：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'------------------
'-- OS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

    '------------------
    '  the(i) <= angle1_tg_2  :  S(tg7_2) = 0   the_1 < Phi_2 [amgle of conected point(arc1 & arc2)]
    '-------------------
  If Phi_2 = Phi_00 Then
       For I = 0 To div_n
         xmo(I) = 0
         ymo(I) = 0
         xmi(I) = 0
         ymi(I) = 0
       Next I
         xg_a_tmp = 0
         yg_a_tmp = 0
         Area_tmp = 0

  Else

    '------------------------------------------------
    '  1) outer arc2 : inner oli groove arc1
    '------------------------------------------------
        div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

        For I = 0 To div_n
            phi1_v(I) = Phi_00 + div_phi * I
        ' Formura arc outer side
          xmo(I) = (r2_oilgroove - t2_oilgroove) * Cos(phi1_v(I)) + x2_oilgroove_c
          ymo(I) = (r2_oilgroove - t2_oilgroove) * Sin(phi1_v(I)) + y2_oilgroove_c

        Next I

      ' start Point of outer arc
          x_tmp(5) = xmo(0) '+ dx     'on FS xy-cordinate
          y_tmp(5) = ymo(0) '+ dy     'on FS xy-cordinate
          q_tmp(5) = Phi_00

      ' end Point of outer arc
          x_tmp(6) = xmo(div_n) '+ dx     'on FS xy-cordinate
          y_tmp(6) = ymo(div_n) '+ dy     'on FS xy-cordinate
          q_tmp(6) = Phi_2

    '------------------------------------------------
    '  2) innner arc2 : FS inner curve
    '------------------------------------------------
        div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

        For I = 0 To div_n
            phi1_v(I) = Phi_0 + div_phi * I

          'Dummy (xmi,ymi) = FS inner curve :Formura (8) fi
            Call Wp_xyfi(phi1_v(I))
                xmi(I) = Wp_xfi          '
                ymi(I) = Wp_yfi
        Next I

      ' start Point of innner arc2
          x_tmp(7) = xmi(0)           'on FS xy-cordinate
          y_tmp(7) = ymi(0)           'on FS xy-cordinate
          q_tmp(7) = Phi_0            'on FS Wrap xy-cordinate

      ' end Point of innner arc2
          x_tmp(8) = xmi(div_n)       'on FS xy-cordinate
          y_tmp(8) = ymi(div_n)       'on FS xy-cordinate
          q_tmp(8) = Phi_1            'on FS Wrap xy-cordinate

    '------------------------------------------------
    '  3) Area and Gravity center
    '------------------------------------------------
         For I = 1 To div_n

        ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

            del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                            - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
            Smi_tmp = Smi_tmp + del_Smi

            xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
            yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

            Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
            Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

        ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

            del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                            - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
            Smo_tmp = Smo_tmp + del_Smo

            xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
            yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

            Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
            Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

        Next I

    '---- OS wrap Area
            Area_A = (Smo_tmp + Smi_tmp)

    '---  area geometrical moment of area Sgx, Sgy 総和

            Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp
            Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp

    '--- center of area

            xg_a_tmp = (Sgy_Area_A) / (Area_A)    '+ dx
            yg_a_tmp = (Sgx_Area_A) / (Area_A)    '+ dy
            Area_tmp = Area_A

  End If

'■■Label point
end_Label_tg7_2:

''--- ' for debug of drowing line
'        For i = 0 To div_n
'            xmo2(i) = xmo(i)
'            ymo2(i) = ymo(i)
'            xmi2(i) = xmi(i)
'            ymi2(i) = ymi(i)
'        Next i
'
'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I)  '+ dx      ' FS-xy
               y_out(I) = ymo(I)  '+ dy      ' FS-xy
               x_in(I) = xmi(I)   '+ dx      ' FS-xy
               y_in(I) = ymi(I)   '+ dy      ' FS-xy
         Next I

End Sub


'======================================================== 【      】
'    Get_Gravity_Center_Oilgroove_FS_tg8_1
'     [tg8-1] : [area of Between oil groove inner arc1 and FS_in ]
'     left area2 of between oilgroove and FS
'
'========================================================

Public Sub Get_Gravity_Center_Oilgroove_FS_tg8_1(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'------------------------------------------------
'  1) outer arc1 : inner oli groove arc1  (use format of FS outer line)
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xfo(I) = (r1_oilgroove - t1_oilgroove) * Cos(phi1_v(I)) + x1_oilgroove_c
      yfo(I) = (r1_oilgroove - t1_oilgroove) * Sin(phi1_v(I)) + y1_oilgroove_c

    Next I


'------------------------------------------------
'  2) innner arc : FS inner curve
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura (8) fi
        Call Wp_xyfi(phi1_v(I))
            xfi(I) = Wp_xfi          '
            yfi(I) = Wp_yfi
    Next I

'------------------------------------------------
'  3) Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp

        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A) ' + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) ' + dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I)  '+ dx      ' FS-xy
               y_out(I) = yfo(I)  '+ dy      ' FS-xy
               x_in(I) = xfi(I)   '+ dx      ' FS-xy
               y_in(I) = yfi(I)   '+ dy      ' FS-xy
         Next I

End Sub


'======================================================== 【      】
'    Get_Gravity_Center_Oilgroove_FS_tg8_2
'     [tg8-2] : [area of Between oil groove inner arc1 and FS_in ]
'     left area2 of between oilgroove and FS
'
'========================================================

Public Sub Get_Gravity_Center_Oilgroove_FS_tg8_2(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)
    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　OS Wrap  ：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'------------------
'-- OS Wrap
'    the 1st part of OS wrap from outer end of the wrap ]
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0


'------------------------------------------------
'  1) outer arc2 : inner oli groove arc2
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xmo(I) = (r2_oilgroove - t2_oilgroove) * Cos(phi1_v(I)) + x2_oilgroove_c
      ymo(I) = (r2_oilgroove - t2_oilgroove) * Sin(phi1_v(I)) + y2_oilgroove_c

    Next I

'------------------------------------------------
'  2) innner arc2 : FS inner curve
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I

      'Dummy (xmi,ymi) = FS inner curve :Formura (8) fi
        Call Wp_xyfi(phi1_v(I))
            xmi(I) = Wp_xfi          '
            ymi(I) = Wp_yfi
    Next I

'------------------------------------------------
'  3) Area and Gravity center
'------------------------------------------------
     For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                        - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

        del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                        - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- OS wrap Area
        Area_A = (Smo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A)    '+ dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A)    '+ dy
        Area_tmp = Area_A


'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I)  '+ dx      ' FS-xy
               y_out(I) = ymo(I)  '+ dy      ' FS-xy
               x_in(I) = xmi(I)   '+ dx      ' FS-xy
               y_in(I) = ymi(I)   '+ dy      ' FS-xy
         Next I

End Sub


'======================================================== 【      】
'    Get_Gravity_Center_Oilgroove_FS_tg8_3
'     [tg8-2] : [area of Between oil groove inner arc1 and FS_in ]
'     left area2 of between oilgroove and FS
'
'========================================================
Public Sub Get_Gravity_Center_Oilgroove_FS_tg8_3(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0


'------------------------------------------------
'  1) outer arc2 : inner oli groove arc2
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xfo(I) = (r2_oilgroove - t2_oilgroove) * Cos(phi1_v(I)) + x2_oilgroove_c
      yfo(I) = (r2_oilgroove - t2_oilgroove) * Sin(phi1_v(I)) + y2_oilgroove_c

    Next I


'------------------------------------------------
'  2) inner arc : inlet outer curve
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divided angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I

    ' on FS xy-cordinate
      xfi(I) = r_Rfi_c * Cos(phi1_v(I)) + x_Rfi_c   'on FS xy-cordinate
      yfi(I) = r_Rfi_c * Sin(phi1_v(I)) + y_Rfi_c   'on FS xy-cordinate

    Next I


'------------------------------------------------
'  3) Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp

        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A) ' + dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) ' + dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I)  '+ dx      ' FS-xy
               y_out(I) = yfo(I)  '+ dy      ' FS-xy
               x_in(I) = xfi(I)   '+ dx      ' FS-xy
               y_in(I) = yfi(I)   '+ dy      ' FS-xy
         Next I
'
''--- ' for debug of drowing line
'    For i = 0 To div_n
'      xfo3(i) = xfo(i)
'      yfo3(i) = yfo(i)
'      xfi3(i) = xfi(i)
'      yfi3(i) = yfi(i)
'    Next i

End Sub


'======================================================== 【      】
'    Get_Gravity_Center_FS_wrapEnd_tg5
'     [tg5 : FS Wrap Tip area of Between Suction inlet and FS_in ]
'
'========================================================
Public Sub Get_Gravity_Center_FS_wrapEnd_tg5(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'------------------------------------------------
'  1)  outer arc : Suction inlet inner curve
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divided angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc on fi side  on FS xy-cordinate
      xfo(I) = r_Rfo_c * Cos(phi1_v(I)) + x_Rfo_c '- dx
      yfo(I) = r_Rfo_c * Sin(phi1_v(I)) + y_Rfo_c '- dy

    Next I

'------------------------------------------------
'  2) inner arc : FS inner curve
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divided angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura (8) fi   on FS xy-cordinate
        xfi(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) + dx
        yfi(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) + dy

    Next I

'------------------------------------------------
'  3) Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp

        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area     'on FS xy-cordinate

        xg_a_tmp = (Sgy_Area_A) / (Area_A) '+ dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) '+ dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I)    ' + dx
               y_out(I) = yfo(I)    '  + dy

               x_in(I) = xfi(I)    '  + dx
               y_in(I) = yfi(I)    '  + dy
         Next I


End Sub


'======================================================== 【      】
'    Get_Gravity_Center_FS_wrapEnd_tg5_2
'     [tg5 : FS Wrap Tip area of Between Suction inlet and FS_in ]
'  tg5_1) Suction inlet : arc1 + arc2(Involute)  Rotation angle: the_1 = CrossPoint(23) to 360deg
'  tg5_2) Suction inlet : arc1 only              Rotation angle: the_1 = 0 to CrossPoint(23)
'
'========================================================
Public Sub Get_Gravity_Center_FS_wrapEnd_tg5_2(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)


'------------------------------------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0


'------------------------------------------------
'  1)  outer arc : Suction inlet inner curve
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divided angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura (7) fo   on FS Wrap-xy-cordinate
        xmo(I) = -a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) + dx
        ymo(I) = -a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) + dy

    Next I

'------------------------------------------------
'  2) inner arc : FS inner curve
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divided angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura (8) fi   on FS Wrap-xy-cordinate
        xmi(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) + dx
        ymi(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) + dy

    Next I

'------------------------------------------------
'  3) Area  and Gravity center
'------------------------------------------------
     For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                        - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

        del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                        - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- total Area
        Area_A = (Smo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp


'--- center of area         'on FS xy-cordinate

        xg_a_tmp = (Sgy_Area_A) / (Area_A) '+ dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) '+ dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I)  '+ dx      ' FS-xy
               y_out(I) = ymo(I)  '+ dy      ' FS-xy
               x_in(I) = xmi(I)   '+ dx      ' FS-xy
               y_in(I) = ymi(I)   '+ dy      ' FS-xy
         Next I

End Sub



'======================================================== 【      】
'    Get_Gravity_Center_FS_wrapEnd_tg5_3
'     [tg5 : FS Wrap Tip area of Between Suction inlet and FS_in ]
'  tg5_1) Suction inlet : arc1 + arc2(Involute)  Rotation angle: the_1 = CrossPoint(23) to 360deg
'  tg5_2) Suction inlet : arc1 only              Rotation angle: the_1 = 0 to CrossPoint(23)
'
'========================================================
Public Sub Get_Gravity_Center_FS_wrapEnd_tg5_3(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

'    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
'    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
'    Dim Area_A As Double:                                      ' FS Wrap 面積
'
'' A : center of Chamber fiqure
'    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
'    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double
'
'    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
'    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
'    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet
'
'    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
'    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)

'
''------------------
''    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
''------------------------------------------------
'    Smi_tmp = 0
'    Smo_tmp = 0
'      xg_mi_tmp = 0:   yg_mi_tmp = 0
'      xg_mo_tmp = 0:   yg_mo_tmp = 0
'      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
'      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0

'------------------------------------------------
'  1)  outer arc : Suction inlet inner curve
'------------------------------------------------

        For I = 0 To div_n
            xmo(I) = 0
            ymo(I) = 0
        Next I

'------------------------------------------------
'  2) inner arc : FS inner curve
'------------------------------------------------

        For I = 0 To div_n
            xmi(I) = 0
            ymi(I) = 0
        Next I

'--- center of area         'on FS xy-cordinate

        xg_a_tmp = 0
        yg_a_tmp = 0
        Area_tmp = 0

End Sub



'======================================================== 【      】
'    Get_Gravity_Center_FS_OS_tg6
'     [tg6 : FS Wrap Tip area of Between Suction inlet and FS_in ]
'
'========================================================
Public Sub Get_Gravity_Center_FS_OS_tg6(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'------------------------------------------------
'  1) outer arc1 : OS-Plate  (use format of FS outer line)
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xfo(I) = OS_dia / 2 * Cos(phi1_v(I)) + Ro * Cos(the_1)     'on FS  xy-cordinate
      yfo(I) = OS_dia / 2 * Sin(phi1_v(I)) + Ro * Sin(the_1)     'on FS  xy-cordinate

    Next I

'
'------------------------------------------------
'  2) inner arc : FS inner curve
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divided angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura (8) fi
    ' on FS xy-cordinate
        xfi(I) = a * phi1_v(I) ^ k * Cos(phi1_v(I) - qq) + g1 * Cos(phi1_v(I) - qq - Atn(k / phi1_v(I))) + dx
        yfi(I) = a * phi1_v(I) ^ k * Sin(phi1_v(I) - qq) + g1 * Sin(phi1_v(I) - qq - Atn(k / phi1_v(I))) + dy

    Next I


'------------------------------------------------
'  3) Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp
        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A)
        yg_a_tmp = (Sgx_Area_A) / (Area_A)
        Area_tmp = Area_A

' check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I)      ' FS-xy
               y_out(I) = yfo(I)      ' FS-xy
               x_in(I) = xfi(I)      ' FS-xy
               y_in(I) = yfi(I)      ' FS-xy
         Next I

End Sub



'======================================================== 【      】
'    Get_Gravity_Center_FS_OS_Plate
'     [tg6 : FS Wrap Tip area of Between Suction inlet and FS_in ]
'
'========================================================
Public Sub Get_Gravity_Center_OS_plate_seal(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Smi_tmp As Double:       Dim Smo_tmp As Double        ' 総和面積
    Dim del_Smi As Double:        Dim del_Smo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_mi_tmp As Double:     Dim del_yg_mi_tmp As Double    ' 分割部 図心
    Dim del_xg_mo_tmp As Double:     Dim del_yg_mo_tmp As Double

    Dim Sgx_mi_tmp As Double:        Dim Sgy_mi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_mo_tmp As Double:        Dim Sgy_mo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_mi_tmp As Double:         Dim yg_mi_tmp As Double        ' 分割Concave部 図心
    Dim xg_mo_tmp As Double:         Dim yg_mo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xmi(div_n):   ReDim ymi(div_n)
    ReDim xmo(div_n):   ReDim ymo(div_n)


'------------------------------------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Smi_tmp = 0
    Smo_tmp = 0
      xg_mi_tmp = 0:   yg_mi_tmp = 0
      xg_mo_tmp = 0:   yg_mo_tmp = 0
      Sgx_mi_tmp = 0:   Sgy_mi_tmp = 0
      Sgx_mo_tmp = 0:   Sgy_mo_tmp = 0


'------------------------------------------------
'  1) outer arc1 : OS-Plate  (use format of FS outer line)
'------------------------------------------------
    div_phi = 2 * pi / div_n  ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xmo(I) = OS_dia / 2 * Cos(phi1_v(I)) + Ro * Cos(the_1)     'on FS  xy-cordinate
      ymo(I) = OS_dia / 2 * Sin(phi1_v(I)) + Ro * Sin(the_1)     'on FS  xy-cordinate

    Next I

'------------------------------------------------
'  2) inner arc : FS inner curve
'------------------------------------------------
    div_phi = 2 * pi / div_n  ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I
    ' Formura arc outer side
      xmi(I) = OS_seal / 2 * Cos(phi1_v(I))   'on FS  xy-cordinate
      ymi(I) = OS_seal / 2 * Sin(phi1_v(I))     'on FS  xy-cordinate

    Next I

'------------------------------------------------
'  3) Area  and Gravity center
'------------------------------------------------
     For I = 1 To div_n

    ' inner triangle Mesh [ mi(i)-mi(i-1)-mo(i-1)]  Wrap area :

        del_Smi = Abs((xmi(I) - xmo(I - 1)) * (ymi(I - 1) - ymo(I - 1)) _
                        - (xmi(I - 1) - xmo(I - 1)) * (ymi(I) - ymo(I - 1))) / 2
        Smi_tmp = Smi_tmp + del_Smi

        xg_mi_tmp = (xmi(I) + xmi(I - 1) + xmo(I - 1)) / 3
        yg_mi_tmp = (ymi(I) + ymi(I - 1) + ymo(I - 1)) / 3

        Sgx_mi_tmp = Sgx_mi_tmp + (del_Smi) * yg_mi_tmp
        Sgy_mi_tmp = Sgy_mi_tmp + (del_Smi) * xg_mi_tmp

    ' outer triangle[ mo(i)-mo(i-1)-mi(i)]

        del_Smo = Abs((xmo(I) - xmi(I)) * (ymo(I - 1) - ymi(I)) _
                        - (xmo(I - 1) - xmi(I)) * (ymo(I) - ymi(I))) / 2
        Smo_tmp = Smo_tmp + del_Smo

        xg_mo_tmp = (xmo(I) + xmo(I - 1) + xmi(I)) / 3
        yg_mo_tmp = (ymo(I) + ymo(I - 1) + ymi(I)) / 3

        Sgx_mo_tmp = Sgx_mo_tmp + (del_Smo) * yg_mo_tmp
        Sgy_mo_tmp = Sgy_mo_tmp + (del_Smo) * xg_mo_tmp

    Next I

'---- total Area
        Area_A = (Smo_tmp + Smi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_mo_tmp + Sgx_mi_tmp
        Sgy_Area_A = Sgy_mo_tmp + Sgy_mi_tmp


'--- center of area         'on FS xy-cordinate

        xg_a_tmp = (Sgy_Area_A) / (Area_A) '+ dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A) '+ dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xmo(I)  '+ dx      ' FS-xy
               y_out(I) = ymo(I)  '+ dy      ' FS-xy
               x_in(I) = xmi(I)   '+ dx      ' FS-xy
               y_in(I) = ymi(I)   '+ dy      ' FS-xy
         Next I

End Sub




'======================================================== 【      】
'    Get_Gravity_Center_FS_OS_tg9
'     [tg9 : FS area of Between Suction inlet and OS Plate]
'
'========================================================
Public Sub Get_Gravity_Center_FS_OS_tg9(ByVal Phi_0 As Double, ByVal Phi_1 As Double, _
                         ByVal Phi_00 As Double, ByVal Phi_2 As Double, ByVal the_1 As Double)

    Dim I As Long, J As Long
    Dim div_phi As Double

    Dim Sfi_tmp As Double:       Dim Sfo_tmp As Double        ' 総和面積
    Dim del_Sfi As Double:        Dim del_Sfo As Double         ' 分割面積
    Dim Area_A As Double:                                      ' FS Wrap 面積

' A : center of Chamber fiqure
    Dim del_xg_fi_tmp As Double:     Dim del_yg_fi_tmp As Double    ' 分割部 図心
    Dim del_xg_fo_tmp As Double:     Dim del_yg_fo_tmp As Double

    Dim Sgx_fi_tmp As Double:        Dim Sgy_fi_tmp As Double       ' 分割部 1次Momet
    Dim Sgx_fo_tmp As Double:        Dim Sgy_fo_tmp As Double
    Dim Sgx_Area_A As Double:         Dim Sgy_Area_A As Double        ' 1次Momet

    Dim xg_fi_tmp As Double:         Dim yg_fi_tmp As Double        ' 分割Concave部 図心
    Dim xg_fo_tmp As Double:         Dim yg_fo_tmp As Double

'------------------
'　FS Wrap　　：配列設定
'------------------

    ReDim phi1_v(div_n)
    ReDim xfi(div_n):   ReDim yfi(div_n)
    ReDim xfo(div_n):   ReDim yfo(div_n)


'------------------
'    各動径間の分割面積総和  ：del_Sfi() , del_Smo()
'------------------------------------------------
    Sfi_tmp = 0
    Sfo_tmp = 0
      xg_fi_tmp = 0:   yg_fi_tmp = 0
      xg_fo_tmp = 0:   yg_fo_tmp = 0
      Sgx_fi_tmp = 0:   Sgy_fi_tmp = 0
      Sgx_fo_tmp = 0:   Sgy_fo_tmp = 0

'------------------------------------------------
'  1) outer arc1 : OS-Plate  (use format of FS outer line)
'------------------------------------------------
    div_phi = (Phi_2 - Phi_00) / div_n    ' Divied angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_00 + div_phi * I
    ' Formura arc outer side
      xfo(I) = OS_dia / 2 * Cos(phi1_v(I)) + Ro * Cos(the_1)     'on FS  xy-cordinate
      yfo(I) = OS_dia / 2 * Sin(phi1_v(I)) + Ro * Sin(the_1)     'on FS  xy-cordinate

    Next I

'------------------------------------------------
'  2) inner arc : inlet outer curve
'------------------------------------------------
    div_phi = (Phi_1 - Phi_0) / div_n    ' Divided angle　分割の角度幅

    For I = 0 To div_n
        phi1_v(I) = Phi_0 + div_phi * I

    ' on FS xy-cordinate
      xfi(I) = r_Rfi_c * Cos(phi1_v(I)) + x_Rfi_c   'on FS xy-cordinate
      yfi(I) = r_Rfi_c * Sin(phi1_v(I)) + y_Rfi_c   'on FS xy-cordinate

    Next I

'------------------------------------------------
'  3) Area  and Gravity center
'------------------------------------------------
    For I = 1 To div_n

    ' inner triangle Mesh [ fi(i)-fi(i-1)-fo(i-1)]  Wrap area :

        del_Sfi = Abs((xfi(I) - xfo(I - 1)) * (yfi(I - 1) - yfo(I - 1)) _
                        - (xfi(I - 1) - xfo(I - 1)) * (yfi(I) - yfo(I - 1))) / 2
        Sfi_tmp = Sfi_tmp + del_Sfi

        xg_fi_tmp = (xfi(I) + xfi(I - 1) + xfo(I - 1)) / 3
        yg_fi_tmp = (yfi(I) + yfi(I - 1) + yfo(I - 1)) / 3

        Sgx_fi_tmp = Sgx_fi_tmp + (del_Sfi) * yg_fi_tmp
        Sgy_fi_tmp = Sgy_fi_tmp + (del_Sfi) * xg_fi_tmp

    ' outer triangle[ fo(i)-fo(i-1)-fi(i)]

        del_Sfo = Abs((xfo(I) - xfi(I)) * (yfo(I - 1) - yfi(I)) _
                        - (xfo(I - 1) - xfi(I)) * (yfo(I) - yfi(I))) / 2
        Sfo_tmp = Sfo_tmp + del_Sfo

        xg_fo_tmp = (xfo(I) + xfo(I - 1) + xfi(I)) / 3
        yg_fo_tmp = (yfo(I) + yfo(I - 1) + yfi(I)) / 3

        Sgx_fo_tmp = Sgx_fo_tmp + (del_Sfo) * yg_fo_tmp
        Sgy_fo_tmp = Sgy_fo_tmp + (del_Sfo) * xg_fo_tmp

    Next I

'---- FS wrap Area
        Area_A = (Sfo_tmp + Sfi_tmp)

'---  area geometrical moment of area Sgx, Sgy 総和

        Sgx_Area_A = Sgx_fo_tmp + Sgx_fi_tmp

        Sgy_Area_A = Sgy_fo_tmp + Sgy_fi_tmp

'--- center of area

        xg_a_tmp = (Sgy_Area_A) / (Area_A)      '+ dx
        yg_a_tmp = (Sgx_Area_A) / (Area_A)      '+ dy
        Area_tmp = Area_A

'--- check for curve
         For I = 0 To div_n
               x_out(I) = xfo(I)  '+ dx      ' FS-xy
               y_out(I) = yfo(I)  '+ dy      ' FS-xy
               x_in(I) = xfi(I)   '+ dx      ' FS-xy
               y_in(I) = yfi(I)   '+ dy      ' FS-xy
         Next I

End Sub

'======================================================== 【      】
'   < Get_Matrix_A_and_C >
'
'
'========================================================
Public Sub Get_Matrix_A_and_C()

    Dim dmmy_a As Double:    Dim dmmy_b As Double:    Dim dmmy_c As Double:
'    Dim dmmy_x As Double:    Dim dmmy_y As Double:
'    Dim dmmy_q As Double:

    Dim D_the As Double:

    Dim I As Long, J As Long

    ReDim Matrix_A(1 To 8, 1 To 8):
    ReDim Matrix_C(1 To 8):

  I = 0

  D_the = (delta_ky - the_c(I))          '[20180719]

   '** VBA <Calc_GasForce_Fr()>結果
'      Fgc_r = 524.05268175235
      Fgc_r = Fr_AB(J)     ' = 524.05268175235    '**VBA結果 258.74 ***約半分

   '** VBA <Calc_GasForce_Ft()>結果

      Fgc_t = Ft_AB(J)     ' = 4225.06655933097
         R_Fgc_t = Lt_AB(J)   ' = 5.0072963597173     '
         R_Fgc_r = Lr_AB(J)   ' = -0.824002719804529  '**VBA結果 -0.2684

   '** VBA <Calc_GasForce_Fz()>結果
      Fgc_z = Fz_Za(I)     ' = 18269.0094291992
      Fgb_z = Fz_Zb(I)     ' = 24223.165717017
      Fsp_z = Fz_sp(I)     ' = 5939.59008491286

      '  Fsp_z = Fgb_z - Fgc_z - F_mg    ' Fsp=Fb-Fa-Fw

   ' Oso基準の荷重位置(r=)：Ar / (t=)：At
      R_Fgz_r = r_zforce_a(Index_I)       ' Oso基準のFa位置(r=)：Ar
      R_Fgz_t = t_zforce_a(Index_I)       ' Oso基準のFa位置(t=)：At
          '  R_Fgz_r = -0.59094778
          '  R_Fgz_t = -0.107573854

   ' Oso基準の荷重位置(r=)：Br / (t=)：Bt
      R_Fgb_r = r_zforce_b(Index_I)       'Oso基準のFb位置(r=)：Br
      R_Fgb_t = t_zforce_b(Index_I)       'Oso基準のFb位置(t=)：Bt
         ' R_Fgb_r = -0.978663241909462
         ' R_Fgb_t = 3.9305633837589E-16


''------------------
''　Matrix_A(1 to 8,1 to 8)
''------------------
'  ' equation (1)
      Matrix_A(1, 1) = -Sin(D_the) - myu_ky * Cos(D_the)
      Matrix_A(1, 2) = Sin(D_the) - myu_ky * Cos(D_the)
      Matrix_A(1, 3) = 0        '= -(h_pl) * (-Sin(D_the) - myu_ky * Cos(D_the))
      Matrix_A(1, 4) = 0
      Matrix_A(1, 5) = 0
      Matrix_A(1, 6) = -1
      Matrix_A(1, 7) = 0
      Matrix_A(1, 8) = 0
'
'  ' equation (2)
      Matrix_A(2, 1) = Cos(D_the) - myu_ky * Sin(D_the)
      Matrix_A(2, 2) = -Cos(D_the) - myu_ky * Sin(D_the)
      Matrix_A(2, 3) = 0
      Matrix_A(2, 4) = 0
      Matrix_A(2, 5) = 1
      Matrix_A(2, 6) = 0
      Matrix_A(2, 7) = 0
      Matrix_A(2, 8) = 0
'
'
'  ' equation (3)
      Matrix_A(3, 1) = -(h_pl - h_ky / 2) * (Sin(D_the) + myu_ky * Cos(D_the))
      Matrix_A(3, 2) = (h_pl - h_ky / 2) * (Sin(D_the) - myu_ky * Cos(D_the))
      Matrix_A(3, 3) = 0
      Matrix_A(3, 4) = 0
      Matrix_A(3, 5) = 0
      Matrix_A(3, 6) = -Z_eb / 2
      Matrix_A(3, 7) = 0
      Matrix_A(3, 8) = -Fsp_z
'
'
'  ' equation (4)
      Matrix_A(4, 1) = -(h_pl - h_ky / 2) * (Cos(D_the) - myu_ky * Sin(D_the))
      Matrix_A(4, 2) = (h_pl - h_ky / 2) * (Cos(D_the) + myu_ky * Sin(D_the))
      Matrix_A(4, 3) = 0
      Matrix_A(4, 4) = 0
      Matrix_A(4, 5) = -Z_eb / 2
      Matrix_A(4, 6) = 0
      Matrix_A(4, 7) = Fsp_z
      Matrix_A(4, 8) = 0
'
'
'  ' equation (5)
      Matrix_A(5, 1) = -(Ros_F1_oy + Ro * Cos(D_the) + myu_ky * (b_kos / 2))  '? [20180719]
      Matrix_A(5, 2) = -(Ros_F2_oy - Ro * Cos(D_the) + myu_ky * (b_kos / 2))  '? [20180719]

'      Matrix_A(5, 1) = -(R_kos + Ro * Cos(D_the) + myu_ky * (b_kos / 2))     ' [20180719]
'      Matrix_A(5, 2) = -(R_kos - Ro * Cos(D_the) + myu_ky * (b_kos / 2))     ' [20180719]
'      Matrix_A(5, 1) = -R_kos - Ro * Sin(pi / 2 - D_the) + myu_ky * (b_kos / 2)
'      Matrix_A(5, 2) = -R_kos + Ro * Sin(pi / 2 - D_the) - myu_ky * (b_kos / 2)

      Matrix_A(5, 3) = 0
      Matrix_A(5, 4) = 0
      Matrix_A(5, 5) = myu_sb * R_eb      ' = myu_sb *28 / 2   '[20180719]  ' R_eb=30/2
      Matrix_A(5, 6) = 0
      Matrix_A(5, 7) = 0
      Matrix_A(5, 8) = -myu_th * Fsp_z
'
'
'  ' equation (6)
      Matrix_A(6, 1) = -1
      Matrix_A(6, 2) = 1
      Matrix_A(6, 3) = -myu_ky
      Matrix_A(6, 4) = -myu_ky
      Matrix_A(6, 5) = 0
      Matrix_A(6, 6) = 0
      Matrix_A(6, 7) = 0
      Matrix_A(6, 8) = 0
'
'
'  ' equation (7)
      Matrix_A(7, 1) = -myu_ky
      Matrix_A(7, 2) = -myu_ky
      Matrix_A(7, 3) = 1
      Matrix_A(7, 4) = -1
      Matrix_A(7, 5) = 0
      Matrix_A(7, 6) = 0
      Matrix_A(7, 7) = 0
      Matrix_A(7, 8) = 0
'
'
'  ' equation (8)
      Matrix_A(8, 1) = myu_ky * b_kos / 2 - R_kos
      Matrix_A(8, 2) = -(myu_ky * b_kos / 2 + R_kos)
      Matrix_A(8, 3) = myu_ky * b_kmf / 2 + R_kmf
      Matrix_A(8, 4) = -(myu_ky * b_kmf / 2 - R_kmf)
      Matrix_A(8, 5) = 0
      Matrix_A(8, 6) = 0
      Matrix_A(8, 7) = 0
      Matrix_A(1, 8) = 0

'
''------------------
''　Matrix_C()
''------------------
      Matrix_C(1) = -Fmc_r + Fgc_r + 0
      Matrix_C(2) = Fgc_t + myu_th * Fsp_z
      Matrix_C(3) = -Hw / 2 * (Fgc_r + Fs_r) + Fgc_z * R_Fgz_r - Fgb_z * R_Fgb_r + Fmc_r * Z_mg + F_mg * R_Fmg_r
      Matrix_C(4) = Hw / 2 * (Fgc_t) - Fgc_z * R_Fgz_t + Fgb_z * R_Fgb_t - F_mg * R_Fmg_t

      Matrix_C(5) = -R_Fgc_t * Fgc_t + R_Fmg_t * Fmc_r    '  - R_Fgc_r * Fgc_r ' [20180719]

      Matrix_C(6) = -m_or * Ro * (2 * pi * N_rps) ^ 2 * Sin(D_the) / 1000000
      Matrix_C(7) = 0
      Matrix_C(8) = 0


End Sub


'======================================================== 【      】
'   < Get_Matrix_X >
'
'
'========================================================
Public Sub Get_Matrix_X()

    Dim dmmy_a As Double:    Dim dmmy_b As Double:    Dim dmmy_c As Double:
    Dim dmmy_x As Double:    Dim dmmy_y As Double:
    Dim dmmy_q As Double:

    Dim I As Long, J As Long
    Dim I1 As Long, J1 As Long

'    Dim temp_Matrix_X As Variant
'    Dim temp_Matrix_X(1 To 1, 1 To 8) As Variant

    ReDim Matrix_X(1 To 8)


'------------------
' Matrix Data 一括貼付
'
'------------------

   DataSheetName = "MatrixSheet_1"

    Sheets(DataSheetName).Select
'    Sheets(DataSheetName).Cells.ClearContents                    '：Sheet全Cellの数式、文字列をクリア

    Sheets(DataSheetName).Range("B37:I44").ClearContents        '：[A]  指定Cellの数式、文字列をクリア
    Sheets(DataSheetName).Range("L37:L44").ClearContents        '：[C]  指定Cellの数式、文字列をクリア


        With Sheets(DataSheetName)
            .Range("B37:I44").Value _
                = Matrix_A
'                = WorksheetFunction.Transpose(Matrix_A)
        End With

        With Sheets(DataSheetName)
            .Range("L37:L44").Value _
                = WorksheetFunction.Transpose(Matrix_C)
        End With

'     Set S1 = Sheets(DataSheetName)
    ' セル範囲を一気に配列に転記
'        Matrix_X = Sheets(DataSheetName).Range("O37").Resize(8, 1).Value
'        WorksheetFunction.Transpose(Matrix_X) = Sheets(DataSheetName).Range("O37:O44").Value
'
'          temp_Matrix_X = Sheets(DataSheetName).Range("O37:O44")  '.Value
'          Matrix_X = WorksheetFunction.Transpose(temp_Matrix_X)

          Matrix_X(1) = Sheets(DataSheetName).Range("O37:O37")
          Matrix_X(2) = Sheets(DataSheetName).Range("O38:O38")
          Matrix_X(3) = Sheets(DataSheetName).Range("O39:O39")
          Matrix_X(4) = Sheets(DataSheetName).Range("O40:O40")
          Matrix_X(5) = Sheets(DataSheetName).Range("O41:O41")
          Matrix_X(6) = Sheets(DataSheetName).Range("O42:O42")
          Matrix_X(7) = Sheets(DataSheetName).Range("O43:O43")
          Matrix_X(8) = Sheets(DataSheetName).Range("O44:O44")

        With Sheets(DataSheetName)
            .Cells(1, 2).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

'            .Cells(1, 3).Value = Format(Now(), "yyyy/MM/DD")       '　"Date
'            .Cells(1, 4).Value = Format(Now(), "HH:mm:ss")         '　"Date
        End With

        Fk_1 = Matrix_X(1)
        Fk_2 = Matrix_X(2)
        Fk_3 = Matrix_X(3)
        Fk_4 = Matrix_X(4)
        Fsb_t = Matrix_X(5)
        Fsb_r = Matrix_X(6)
        R_Fsp_t = Matrix_X(7)
        R_Fsp_r = Matrix_X(8)

      Torque_s = Fsb_t * Ro
      Fsb_e = Sqr(Fsb_t ^ 2 + Fsb_r ^ 2)
      Moment_e = Fsb_e * myu_sb * R_eb

      Tilting_os = Sqr(Matrix_X(7) ^ 2 + Matrix_X(8) ^ 2)
      Stability_os = Sqr(Matrix_X(7) ^ 2 + Matrix_X(8) ^ 2) * 2 / (OS_dia)


End Sub

'======================================================== 【      】
'   < Get_Matrix_and_results >
'
'
'========================================================
Public Sub Get_Matrix_and_results()

   Dim dmmy_a As Double:    Dim dmmy_b As Double:    Dim dmmy_c As Double:
   Dim dmmy_M As Double:
   Dim delta_M As Double:

   Dim D_the As Double:
   Dim I As Long, J As Long

   '---------------------------------------
   '　 Matrix term
       '  Fk_1 = Matrix_X(1)
       '  Fsb_t = Matrix_X(5)           '  Fsb_r = Matrix_X(6)
       '  R_Fsp_t = Matrix_X(7)         '  R_Fsp_r = Matrix_X(8)
         '  Torque_s = Fsb_t * Ro
         '  Fsb_e = Sqr(Fsb_t ^ 2 + Fsb_r ^ 2)
         '  Moment_e = Fsb_e * myu_sb * R_eb
   '　 Matrix term
       '  Matrix_A(5, 5) = myu_sb * R_eb
       '  Matrix_C(5) = -R_Fgc_t * Fgc_t + R_Fmg_t * Fmc_r

    J = 0

          Debug.Print "Initial Moment_e=" & Moment_e
          Debug.Print "  Fsb_t="; Format(Matrix_X(5), "####.####"); Tab(2); _
                      "Fsb_r="; Format(Matrix_X(6), "####.####")
'        Stop
    Do
        dmmy_M = Moment_e
        Matrix_A(5, 5) = 0
        Matrix_C(5) = -R_Fgc_t * Fgc_t + R_Fmg_t * Fmc_r - Moment_e

      '---------------------------------------
      '　 Calculation of Matrix
        Call Get_Matrix_X
        '  Fsb_t = Matrix_X(5)           '  Fsb_r = Matrix_X(6)

        J = J + 1
        delta_M = dmmy_M - Moment_e

          Debug.Print "delta_M="; Format(delta_M, "####.####"); Tab(2); _
                              "dmmy_M="; Format(dmmy_M, "####.####"); "  "; _
                              "Moment_e="; Format(Moment_e, "####.####")
          Debug.Print "Fsb_t="; Format(Matrix_X(5), "####.####"); Tab(2); _
                              "Fsb_r="; Format(Matrix_X(6), "####.####")
'        Stop

    Loop While delta_M > dmmy_M * 0.0001


End Sub


'======================================================== 【      】
'   < Calc_Gravity_Center_Mass_OS >
'     Public Sub Calc_Gravity_Center_Mass_OS(the_1 As Double)
'
'========================================================
Public Sub Calc_Gravity_Center_Mass_OS()       '(the_1 As Double)

    Dim tmp_out_0 As Double:    Dim tmp_out_1 As Double:
    Dim tmp_in_0 As Double:     Dim tmp_in_1 As Double:

    Dim tmp_N As Double:
    Dim tmp_q_mr As Double:  Dim tmp_q_mr2 As Double:
    Dim tmp_Fc As Double:

    Dim tmp_Area As Double:    Dim tmp_Area2 As Double:
    Dim tmp_Mgx As Double:      Dim tmp_Mgy As Double:      Dim tmp_Mgz As Double

    Dim tmp_V(9) As Double:    Dim tmp_Mg(9) As Double:
    Dim tmp_Gx(9) As Double:   Dim tmp_Gy(9) As Double:   Dim tmp_Gz(9) As Double:
    Dim tmp_Gx_c As Double:     Dim tmp_Gy_c As Double

    Dim I As Long, J As Long
    Dim the_1 As Double:
    Dim Curve_name As String

    Dim I1 As Long, J1 As Long
    Dim Imax As Long, Jmax As Long


'---------------------------------------
'-- Data 貼付 初期設定

  '     貼付先の先頭セルの、行と列 (i1, j1)
  '       ex.) i1 = 4    ' Range("A4").row    "4"　行
  '       ex.) j1 = 2    ' Range("B1").Column "B"　列

            I1 = Range("EV4").Row
            J1 = Range("EV4").Column
            Imax = dw_n                           ' Data array length of raws
            Jmax = 5                              ' Data array width of collums

          ReDim result_0(Jmax, Imax)

  '-- Curve Name

       DataSheetName = "DataSheet_6"
       With Sheets(DataSheetName)
            .Cells(3, J1 + 0).Value = "R_Fmg_r"   ' 重心位置　r方向
            .Cells(3, J1 + 1).Value = "R_Fmg_t"   ' 重心位置　t方向
            .Cells(3, J1 + 2).Value = "tmp_Gy(0)"      ' 重心位置
            .Cells(3, J1 + 3).Value = "tmp_Gy(0)"
            .Cells(3, J1 + 4).Value = "tmp_q_mr"
            .Cells(3, J1 + 5).Value = "tmp_q_mr2"

      End With

'---------------------------------------
'   OS Gravity_Center　　　　** OS-plate-xy座標　(含む+dxdy)
'---------------------------------------

      ''------------
      '
      '      '[tp5 OS]
      '      tmp_in_0 = Phi_c_fi(3, Index_I)        ' see <Calc_Gravity_Center_wrap>'[tp1 OS]   '15 / 180 * pi
      '      tmp_in_1 = Phi_c_fi(2, Index_I)
      '      tmp_out_0 = Phi_c_fi(2, Index_I)   ' see <>'[tp1 OS]   outer Wrap Wall
      '      tmp_out_1 = Phi_c_fi(1, Index_I)
      '
      '
      '       '[tp1 OS]
      '      tmp_in_0 = Wrap_Start_angle_min(3)       ' see <Calc_Gravity_Center_wrap>'[tp1 OS]   '15 / 180 * pi
      '      tmp_in_1 = Phi_c_fi(6, Index_I)
      '      tmp_out_0 = Wrap_Start_angle_min(4)  ' see <>'[tp1 OS]   outer Wrap Wall
      '      tmp_out_1 = Phi_c_fi(5, Index_I)

      ''------------

'' OS(1) wrap part
            tmp_in_0 = OS_in_srt_0         ' see <Calc_Gravity_Center_wrap>'[tp1 OS]   '15 / 180 * pi
            tmp_in_1 = OS_in_end
            tmp_out_0 = OS_out_srt_0      ' see <>'[tp1 OS]   outer Wrap Wall
            tmp_out_1 = OS_out_end

            Index_I = 0
            the_1 = the(Index_I) - qq

                Call Get_Gravity_Center_OS_wrap_all_div2(tmp_in_0, tmp_in_1, tmp_out_0, tmp_out_1) ', the_1)
                  ' Call Get_Gravity_Center_OS_wrap_all(tmp_in_0, tmp_in_1, tmp_out_0, tmp_out_1) ', the_1)
                  '**in,out線, xg_a_tmp,yg_a_tmp 旋廻無

                  tmp_Gx(1) = xg_a_tmp ' + Ro * Cos(the_1)
                  tmp_Gy(1) = yg_a_tmp ' + Ro * Sin(the_1)
                  tmp_Area = Area_tmp

                '--- check for curve
'                   For I = 0 To UBound(x_out)     ' div_n
'                         x_out(I) = x_out(I) + Ro * Cos(the_1)     ' FS-xy 基準
'                         y_out(I) = y_out(I) + Ro * Sin(the_1)     ' FS-xy 基準
'                         x_in(I) = x_in(I) + Ro * Cos(the_1)       ' FS-xy 基準
'                         y_in(I) = y_in(I) + Ro * Sin(the_1)       ' FS-xy 基準
'                   Next I

'                '--Dara Paste -
'                  Curve_name = "OS wrap_":    ' DataSheetName = "DataSheet_2"
'                    Call Paste_curve_data_Num(30, Curve_name & "in", x_in(), y_in(), DataSheetName)
'                    Call Paste_curve_data_Num(31, Curve_name & "out", x_out(), y_out(), DataSheetName)
'                  ' Stop

'            ' 配列数を元に戻す
'              ReDim x_out(div_n):     ReDim y_out(div_n):
'              ReDim x_in(div_n):      ReDim y_in(div_n):

'
' OS(1) wrap part
'
'        tmp_Area = Sg_m(1, 0) + Sg_m(2, 0) + Sg_m(3, 0) + Sg_m(4, 0) + Sg_m(5, 0)
'
'        tmp_Gx(1) = (Sg_m(1, 0) * xg_m(1, 0) + Sg_m(2, 0) * xg_m(2, 0) + Sg_m(3, 0) * xg_m(3, 0) _
'                                + Sg_m(4, 0) * xg_m(4, 0) + Sg_m(5, 0) * xg_m(5, 0)) / tmp_Area
'
'        tmp_Gy(1) = (Sg_m(1, 0) * yg_m(1, 0) + Sg_m(2, 0) * yg_m(2, 0) + Sg_m(3, 0) * yg_m(3, 0) _
'                                + Sg_m(4, 0) * yg_m(4, 0) + Sg_m(5, 0) * yg_m(5, 0)) / tmp_Area

' OS(1) wrap part total

      tmp_V(1) = tmp_Area * Hw
      tmp_Mg(1) = tmp_V(1) * dense_os / 1000
      tmp_Gz(1) = Hw / 2

      tmp_Gx(1) = tmp_Gx(1) - Ro * Cos(the_1)    '+ dx     on OS-xy
      tmp_Gy(1) = tmp_Gy(1) - Ro * Sin(the_1)    '+ dy

' OS(2) plate part

      tmp_V(2) = pi * OS_dia ^ 2 / 4 * h_pl
      tmp_Mg(2) = pi * OS_dia ^ 2 / 4 * h_pl * dense_os / 1000
      tmp_Gx(2) = 0     'Ro * Cos(the_1)
      tmp_Gy(2) = 0     'Ro * Sin(the_1)
      tmp_Gz(2) = -h_pl / 2


' OS(3) key parts(-)

      tmp_V(3) = -(h_ky * b_kos * L_kcv) * 2
      tmp_Mg(3) = -(h_ky * b_kos * L_kcv) * 2 * dense_os / 1000 '[2018/7/19] +-
      tmp_Gx(3) = 0     'Ro * Cos(the_1)
      tmp_Gy(3) = 0     'Ro * Sin(the_1)
      tmp_Gz(3) = (h_pl - L_kcv / 2)                            '[2018/7/19] +-

' OS(4) boss part
      R_eb = 15

      tmp_V(4) = (R_eb_out ^ 2 - R_eb ^ 2) * pi * L_eb_out
      tmp_Mg(4) = tmp_V(4) * dense_os / 1000
      tmp_Gx(4) = 0     'Ro * Cos(the_1)
      tmp_Gy(4) = 0     'Ro * Sin(the_1)
      tmp_Gz(4) = -(L_eb_out / 2 + h_pl)


'------------
' OS(1-4) total

    tmp_Mgx = 0
    tmp_Mgy = 0
    tmp_Mgz = 0
    tmp_V(0) = 0
    tmp_Mg(0) = 0

    For I = 1 To 4

      tmp_V(0) = tmp_V(0) + tmp_V(I)
      tmp_Mg(0) = tmp_Mg(0) + tmp_Mg(I)
      tmp_Mgx = tmp_Mgx + tmp_Mg(I) * tmp_Gx(I)
      tmp_Mgy = tmp_Mgy + tmp_Mg(I) * tmp_Gy(I)
      tmp_Mgz = tmp_Mgz + tmp_Mg(I) * tmp_Gz(I)

    Next I

      vol_os = tmp_V(0)
      m_os = tmp_Mg(0)    '[g]

      tmp_Gx(0) = tmp_Mgx / m_os         ' OS Plate-xyz
      tmp_Gy(0) = tmp_Mgy / m_os         ' OS Plate-xyz
      tmp_Gz(0) = tmp_Mgz / m_os         ' OS Plate-xyz


'---------------------------------------

  For Index_I = 0 To dw_n
  '   For Index_I = 60 To 80    'dw_n

      the_1 = the(Index_I) - qq
'      tmp_N = 1

  ' each vlue check
        For I = 0 To 4
            R_x_osw(I) = tmp_Gx(I) + Ro * Cos(the_1)  ' FS-xy軸上の重心　x位置
            R_y_osw(I) = tmp_Gy(I) + Ro * Sin(the_1)  ' FS-xy軸上の重心　x位置
            R_z_osw(I) = tmp_Gz(I)
            V_osw(I) = tmp_V(I)
        Next I


  ' total result
      tmp_Gx_c = tmp_Gx(0) + Ro * Cos(the_1) ' R_x_osw(0)   ' FS-xy
      tmp_Gy_c = tmp_Gy(0) + Ro * Sin(the_1) ' R_y_osw(0)   ' FS-xy


'      tmp_q_mr = Atn(tmp_Gy_c / tmp_Gx_c)
'      tmp_q_mr2 = Atn(Sin(the_1) / Cos(the_1))
      tmp_q_mr = Atan2(tmp_Gy_c, tmp_Gx_c)        '
      tmp_q_mr2 = the_1

  '-- 遠心力のモーメントアーム
      R_Fmg_r = Sqr(tmp_Gx_c ^ 2 + tmp_Gy_c ^ 2) * Cos(tmp_q_mr - tmp_q_mr2) - Ro    'OS-xy
      R_Fmg_t = Sqr(tmp_Gx_c ^ 2 + tmp_Gy_c ^ 2) * Sin(tmp_q_mr - tmp_q_mr2) * (-1)  '[20180719]add*(-1)
      Z_mg = tmp_Gz(0)

      Fmc_r = m_os / 1000 * Ro / 1000 * (2 * pi * N_rps) ^ 2      '[N] 遠心力
      F_mg = m_os / 1000 * gravity                                '[N] 重力

  '-- 結果保管
      result_0(0, Index_I) = R_Fmg_r        ' 遠心力のモーメントアーム r方向
      result_0(1, Index_I) = R_Fmg_t        ' 遠心力のモーメントアーム t方向
      result_0(2, Index_I) = tmp_Gx_c       ' 重心位置　軸回転角
      result_0(3, Index_I) = tmp_Gy_c       ' 重心位置
      result_0(4, Index_I) = tmp_q_mr * 180 / pi
      result_0(5, Index_I) = tmp_q_mr2 * 180 / pi

  Next Index_I


  '-- 指定Cellの数式、文字列をクリア

    Sheets(DataSheetName).Range(Cells(I1, J1), Cells(1999, J1 + Jmax)).ClearContents

      With Sheets(DataSheetName)

            .Range(Cells(I1, J1), Cells(I1 + Imax, J1 + Jmax)).Value _
                = WorksheetFunction.Transpose(result_0)
      End With


End Sub


'======================================================== 【      】
'   < Calc_Gravity_Center_Mass_OldhamRing >
'
'========================================================
Public Sub Calc_Gravity_Center_Mass_OldhamRing()

    Dim tmp_Fc As Double:
    Dim tmp_Area As Double:    Dim tmp_Area2 As Double:
    Dim tmp_Mgx As Double:      Dim tmp_Mgy As Double:      Dim tmp_Mgz As Double

    Dim I As Long, J As Long

    Dim tmp_V(3) As Double:    Dim tmp_Mg(3) As Double:
    Dim tmp_Gx(3) As Double:   Dim tmp_Gy(3) As Double:   Dim tmp_Gz(3) As Double:
    Dim tmp_Gx_c As Double:     Dim tmp_Gy_c As Double

    Dim the_1 As Double:      Dim D_the As Double
    Dim Curve_name As String

    Dim I1 As Long, J1 As Long
    Dim Imax As Long, Jmax As Long


'---------------------------------------
'-- Data 貼付 初期設定
  '     貼付先の先頭セルの、行と列 (i1, j1)
  '       ex.) i1 = 4    ' Range("A4").row    "4"　行
  '       ex.) j1 = 2    ' Range("B1").Column "B"　列

            I1 = Range("EV4").Row
            J1 = Range("EV4").Column
            Imax = dw_n                           ' Data array length of raws
            Jmax = 5                              ' Data array width of collums

          ReDim result_0(Jmax, Imax)

  '-- Curve Name

       DataSheetName = "DataSheet_6"

       With Sheets(DataSheetName)
            .Cells(3, J1 + 0).Value = "Ros_F1_oy "   ' 重心位置　r方向
            .Cells(3, J1 + 1).Value = "Ros_F2_oy"   ' 重心位置　t方向
'            .Cells(3, J1 + 2).Value = "Z_mg"      ' 重心位置　Z方向
'            .Cells(3, J1 + 3).Value = "Fmc_r"     ' 遠心力
'            .Cells(3, J1 + 4).Value = "F_mg"      ' 重力
            .Cells(3, J1 + 2).Value = "tmp_Gy(0)"      ' 重心位置
            .Cells(3, J1 + 3).Value = "tmp_Gy(0)"
            .Cells(3, J1 + 4).Value = "the_1"
            .Cells(3, J1 + 5).Value = "D_the"

      End With


'---------------------------------------
'   OldamRing Gravity_Center
'---------------------------------------
      Index_I = 0
      the_1 = the(Index_I) - qq
      D_the = (delta_ky - (the_c(Index_I)))      '[20180719] = (delta_ky - (-the_c(i)))　x軸とkeyの角度

      R_or_in = R_or_in - 0.74

      Ros_F1_oy = R_kos + Ro * Sin(pi / 2 - D_the)  '[20180719] + Ro * Cos(D_the)
      Ros_F2_oy = R_kos - Ro * Sin(pi / 2 - D_the)  '[20180719] - Ro * Cos(D_the)
      Ros_F3_ox = R_kmf
      Ros_F4_ox = R_kmf

'------------
' OldamRing Ring part

      tmp_V(1) = (R_or_out ^ 2 - R_or_in ^ 2) * pi * h_or
      tmp_Mg(1) = tmp_V(1) * dense_or / 1000
      tmp_Gx(1) = 0     ' + Ro * Cos(the_1)
      tmp_Gy(1) = 0     ' + Ro * Sin(the_1)
      tmp_Gz(1) = -(h_pl + h_or / 2)

' Key of OS side

      tmp_V(2) = L_kos * b_kos * h_ky * 2
      tmp_Mg(2) = tmp_V(2) * dense_or / 1000
      tmp_Gx(2) = 0     ' + Ro * Cos(the_1)
      tmp_Gy(2) = 0     ' + Ro * Sin(the_1)
      tmp_Gz(2) = -(h_pl - h_ky / 2)

' Key of MF side

      tmp_V(3) = L_kmf * b_kmf * h_ky * 2
      tmp_Mg(3) = tmp_V(3) * dense_or / 1000
      tmp_Gx(3) = 0     ' + Ro * Cos(the_1)
      tmp_Gy(3) = 0     ' + Ro * Sin(the_1)
      tmp_Gz(3) = -(h_pl + h_or + h_ky / 2)

'------------
' OS total

    tmp_Mgx = 0
    tmp_Mgy = 0
    tmp_Mgz = 0

    For I = 1 To 3
      tmp_V(0) = tmp_V(0) + tmp_V(I)
      tmp_Mg(0) = tmp_Mg(0) + tmp_Mg(I)
      tmp_Mgx = tmp_Mgx + tmp_Mg(I) * tmp_Gx(I)
      tmp_Mgy = tmp_Mgy + tmp_Mg(I) * tmp_Gy(I)
      tmp_Mgz = tmp_Mgz + tmp_Mg(I) * tmp_Gz(I)
    Next I

      vol_or = tmp_V(0)
      m_or = tmp_Mg(0)     '[g]

      tmp_Gx(0) = tmp_Mgx / m_or      ' OS-xy軸とOrdham-xy軸を一致させた時の重心　x位置
      tmp_Gy(0) = tmp_Mgy / m_or      ' OS-xy軸とOrdham-xy軸を一致させた時の重心　x位置
      tmp_Gz(0) = tmp_Mgz / m_or

      Fc_or = m_or / 1000 * Ro / 1000 * (2 * pi * N_rps) ^ 2 * Sin(D_the)    '遠心力

'---------------------------------------
  For Index_I = 0 To dw_n

      the_1 = the(Index_I) - qq

'      D_the = (alpha_ky - the_1)
      D_the = (delta_ky - (the_c(Index_I)))     '[20180719] = (delta_ky - (-the_c(i)))

      Ros_F1_oy = R_kos + Ro * Sin(pi / 2 - D_the)  '[20180719] + Ro * Cos(D_the)
      Ros_F2_oy = R_kos - Ro * Sin(pi / 2 - D_the)  '[20180719] - Ro * Cos(D_the)
'      Ros_F3_ox = R_kmf
'      Ros_F4_ox = R_kmf

   ' each vlue check
        For I = 0 To 3
            R_x_or(I) = tmp_Gx(I) + Ro * Sin(alpha_ky) * Sin(D_the)   ' FS-xy
            R_y_or(I) = tmp_Gy(I) + Ro * Cos(alpha_ky) * Sin(D_the)   ' FS-xy
            R_z_or(I) = tmp_Gz(I)
            V_or(I) = tmp_V(I)
        Next I

  ' total result
      tmp_Gx_c = tmp_Gx(0) + Ro * Sin(alpha_ky) * Sin(D_the)       ' FS-xy　重心位置　軸回転角
      tmp_Gy_c = tmp_Gy(0) + Ro * Cos(alpha_ky) * Sin(D_the)       ' FS-xy　重心位置　軸回転角

  '-- 結果保管
      result_0(0, Index_I) = Ros_F1_oy              ' F!反力のモーメントアーム r方向
      result_0(1, Index_I) = Ros_F2_oy              ' 遠心力のモーメントアーム t方向
      result_0(2, Index_I) = tmp_Gx_c               ' 重心位置　軸回転角
      result_0(3, Index_I) = tmp_Gy_c               ' 重心位置
      result_0(4, Index_I) = the_1 * 180 / pi
      result_0(5, Index_I) = D_the * 180 / pi

  Next Index_I


  '-- 指定Cellの数式、文字列をクリア

    Sheets(DataSheetName).Range(Cells(I1, J1), Cells(1999, J1 + Jmax)).ClearContents

      With Sheets(DataSheetName)

            .Range(Cells(I1, J1), Cells(I1 + Imax, J1 + Jmax)).Value _
                = WorksheetFunction.Transpose(result_0)
      End With

End Sub

'======================================================== 【      】
'   <  change_Wrap_data_to_curve_xw_3  >
'                 xmi(i) = curve_xw(3, i)
'                 ymi(i) = curve_yw(3, i)
'========================================================
'Public Sub change_Wrap_data_to_curve_xw_2()
'       Dim i As Long, j As Long
'            For i = 0 To div_n
'                 xmi(i) = curve_xw(3, i)
'                 ymi(i) = curve_yw(3, i)
'            Next i
'End Sub

'======================================================== 【      】
'   <  change_Wrap_data_to_curve_xw_2  >
'                 curve_xu(i) = curve_xw(1, i)
'                 curve_yu(i) = curve_yw(1, i)
'========================================================
Public Sub change_Wrap_data_to_curve_xw_2(ByVal curve_xu As Double, ByVal curve_yu As Double)
       Dim I As Long, J As Long
            For I = 0 To div_n
                 curve_xu(I) = curve_xw(2, I)
                 curve_yu(I) = curve_yw(2, I)
            Next I
End Sub

'======================================================== 【      】
'   <  change_Wrap_data_to_curve_xw_1  >
'                 curve_xu(i) = curve_xw(1, i)
'                 curve_yu(i) = curve_yw(1, i)
'========================================================
Public Sub change_Wrap_data_to_curve_xw(ByVal int_j As Long, ByVal curve_xu, ByVal curve_yu)
       Dim I As Long        ' int_j As Long
'       Dim curve_xu As Double

            For I = 0 To div_n
                 curve_xw(int_j, I) = curve_xu(I)
                 curve_yw(int_j, I) = curve_yu(I)
            Next I
End Sub



'======================================================== 【      】
'   < Cross point of Arc and Line>
'
'    Get_CrossPoint_arc_on_line(q_tmp, r1_tmp ,x1c_tmp,  y1c_tmp, the_1)
'     input   q1_tmp : angle of line
'             r1_tmp ,x1c_tmp,  y1c_tmp: arc radius, arc center x and y
'     output
'       Resalt of point  ==>  ( x1c_tmp,  y1c_tmp)
'
'========================================================
Public Sub Get_CrossPoint_arc_on_line_xy(ByVal x1_tmp As Double, ByVal y1_tmp As Double, ByVal r1_tmp As Double _
                  , ByVal x1c_tmp As Double, ByVal y1c_tmp As Double, ByVal the_1 As Double)

    Dim dmmy_a As Double:    Dim dmmy_b As Double:    Dim dmmy_c As Double:
    Dim dmmy_x As Double:    Dim dmmy_y As Double:
    Dim dmmy_q As Double:

    Dim I As Long, J As Long
    Dim div_phi As Double

'------------------
'　Cross Point arc and Line
'------------------

  ' arc end point : on FS Wrap-xy-cordinate
        q1_tmp = Atan2(y1_tmp, x1_tmp)

      If q1_tmp = pi / 2 Then
        dmmy_x = 0
        dmmy_y = Sqr(r1_tmp ^ 2 - ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2)) + y1c_tmp

      ElseIf q1_tmp = -pi / 2 Then
        dmmy_x = 0
        dmmy_y = -Sqr(r1_tmp ^ 2 - ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2)) + y1c_tmp

      Else
        dmmy_q = y1_tmp / x1_tmp
        dmmy_a = 1 + dmmy_q ^ 2
        dmmy_b = -2 * (x1c_tmp + dmmy_q * y1c_tmp)
        dmmy_c = ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2) - r1_tmp ^ 2

           '**** =[-b+Root(b^2-4ac)]/(2a) ****
            dmmy_x = (-dmmy_b + Sqr(dmmy_b ^ 2 - 4 * dmmy_a * dmmy_c)) / (2 * dmmy_a)
            dmmy_y = dmmy_q * dmmy_x

          If (x1_tmp * dmmy_x) < 0 Then
           '**** =[-b-Root(b^2-4ac)]/(2a) ****
            dmmy_x = (-dmmy_b - Sqr(dmmy_b ^ 2 - 4 * dmmy_a * dmmy_c)) / (2 * dmmy_a)
            dmmy_y = dmmy_q * dmmy_x
          End If

      End If

        x1_tmp = dmmy_x
        y1_tmp = dmmy_y

End Sub




'======================================================== 【      】
'   < Cross point of Arc and Line>
'
'    Get_CrossPoint_arc_on_line(q_tmp, r1_tmp ,x1c_tmp,  y1c_tmp, the_1)
'     input   q1_tmp : angle of line
'             r1_tmp ,x1c_tmp,  y1c_tmp: arc radius, arc center x and y
'     output
'       Resalt of point  ==>  ( x1c_tmp,  y1c_tmp)
'
'========================================================
Public Sub Get_CrossPoint_arc_on_line(ByVal q1_tmp As Double, ByVal r1_tmp As Double _
                  , ByVal x1c_tmp As Double, ByVal y1c_tmp As Double, ByVal the_1 As Double)

    Dim dmmy_a As Double:    Dim dmmy_b As Double:    Dim dmmy_c As Double:
    Dim dmmy_x As Double:    Dim dmmy_y As Double:
    Dim dmmy_q As Double:

    Dim I As Long, J As Long
    Dim div_phi As Double

'------------------
'　Cross Point arc and Line
'------------------

  ' arc end point : on FS Wrap-xy-cordinate
      If q1_tmp = pi / 2 Then
        dmmy_x = 0
        dmmy_y = Sqr(r1_tmp ^ 2 - ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2)) + y1c_tmp

      ElseIf q1_tmp = -pi / 2 Then
        dmmy_x = 0
        dmmy_y = -Sqr(r1_tmp ^ 2 - ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2)) + y1c_tmp

      Else
        dmmy_q = Tan(q1_tmp)
        dmmy_a = 1 + dmmy_q ^ 2
        dmmy_b = -2 * (x1c_tmp + dmmy_q * y1c_tmp)
        dmmy_c = ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2) - r1_tmp ^ 2

       '**** =[-b+Root(b^2-4ac)]/(2a) ****
        dmmy_x = (-dmmy_b + Sqr(dmmy_b ^ 2 - 4 * dmmy_a * dmmy_c)) / (2 * dmmy_a)
        dmmy_y = dmmy_q * dmmy_x

          If (Tan(q1_tmp) * (dmmy_y / dmmy_x)) < 0 Then
            dmmy_x = (-dmmy_b - Sqr(dmmy_b ^ 2 - 4 * dmmy_a * dmmy_c)) / (2 * dmmy_a)
            dmmy_y = dmmy_q * dmmy_x
          End If

      End If

        x1_tmp = dmmy_x
        y1_tmp = dmmy_y

End Sub


'======================================================== 【      】
'   < Cross point of Arc and Line>
'
'    Get_CrossPoint_arc_on_line(q_tmp, r1_tmp ,x1c_tmp,  y1c_tmp, the_1)
'     input   q1_tmp : angle of line
'             r1_tmp ,x1c_tmp,  y1c_tmp: arc radius, arc center x and y
'     output
'       Resalt of point  ==>  ( x1c_tmp,  y1c_tmp)
'
'========================================================
Public Sub Get_CrossPoint_arc_on_line_2(ByVal q1_tmp As Double, ByVal r1_tmp As Double _
                  , ByVal x1c_tmp As Double, ByVal y1c_tmp As Double, ByVal the_1 As Double)

    Dim dmmy_a As Double:    Dim dmmy_b As Double:    Dim dmmy_c As Double:
    Dim dmmy_x As Double:    Dim dmmy_y As Double:
    Dim dmmy_q As Double:

    Dim I As Long, J As Long
    Dim div_phi As Double

'------------------
'　Cross Point arc and Line
'------------------

  ' arc end point : on FS Wrap-xy-cordinate
      If q1_tmp = pi / 2 Then
        dmmy_x = 0
        dmmy_y = Sqr(r1_tmp ^ 2 - ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2)) + y1c_tmp

      ElseIf q1_tmp = -pi / 2 Then
        dmmy_x = 0
        dmmy_y = -Sqr(r1_tmp ^ 2 - ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2)) + y1c_tmp

      Else
        dmmy_q = Tan(q1_tmp)
        dmmy_a = 1 + dmmy_q ^ 2
        dmmy_b = -2 * (x1c_tmp + dmmy_q * y1c_tmp)
        dmmy_c = ((x1c_tmp) ^ 2 + (y1c_tmp) ^ 2) - r1_tmp ^ 2

       '**** =[-b-Root(b^2-4ac)]/(2a) ****
        dmmy_x = (-dmmy_b - Sqr(dmmy_b ^ 2 - 4 * dmmy_a * dmmy_c)) / (2 * dmmy_a)
        dmmy_y = dmmy_q * dmmy_x

'            dmmy_x = (-dmmy_b + Sqr(dmmy_b ^ 2 - 4 * dmmy_a * dmmy_c)) / (2 * dmmy_a)
'            dmmy_y = dmmy_q * dmmy_x

      End If

        x1_tmp = dmmy_x
        y1_tmp = dmmy_y

End Sub



Public Function Atan2(y As Double, x As Double)   '0 < Atan2 <2π     '-π/2< Atn <π/2
    If x = 0 And y = 0 Then
        Atan2 = 0
    ElseIf x > 0 And y = 0 Then 'θ=0
        Atan2 = 0
    ElseIf x = 0 And y > 0 Then 'θ=90
        Atan2 = pi / 2
    ElseIf x < 0 And y = 0 Then 'θ=180
        Atan2 = pi
    ElseIf x = 0 And y < 0 Then 'θ=270
        Atan2 = pi / 2 * 3
    ElseIf x > 0 And y > 0 Then ' 0<θ<90
        Atan2 = Atn(Abs(y) / Abs(x))
    ElseIf x < 0 And y > 0 Then ' 90<θ<180
        Atan2 = pi - Atn(Abs(y) / Abs(x))
    ElseIf x < 0 And y < 0 Then ' 180<θ<270
        Atan2 = Atn(Abs(y) / Abs(x)) + pi
    ElseIf x > 0 And y < 0 Then ' 270<θ<360
        Atan2 = 2 * pi - Atn(Abs(y) / Abs(x))
    End If

    Atan2 = Atan2          '   0 < Atan2 < 2π
'    Atan2 = Atan2 - pi     ' -π < Atan2 < π

End Function



Public Function Atan3(y As Double, x As Double)   ' -π < Atan3 <π     '-π/2< Atn <π/2
    If x = 0 And y = 0 Then
        Atan3 = 0
    ElseIf x > 0 And y = 0 Then 'θ=0
        Atan3 = 0
    ElseIf x = 0 And y > 0 Then 'θ=90
        Atan3 = pi / 2
    ElseIf x < 0 And y = 0 Then 'θ=180
        Atan3 = pi
    ElseIf x = 0 And y < 0 Then 'θ=270
        Atan3 = -pi / 2
    ElseIf x > 0 And y > 0 Then ' 0<θ<90
        Atan3 = Atn(Abs(y) / Abs(x))
    ElseIf x < 0 And y > 0 Then ' 90<θ<180
        Atan3 = pi - Atn(Abs(y) / Abs(x))
    ElseIf x < 0 And y < 0 Then ' 180<θ<270
        Atan3 = Atn(Abs(y) / Abs(x)) - pi
    ElseIf x > 0 And y < 0 Then ' 270<θ<360
        Atan3 = -Atn(Abs(y) / Abs(x))
    End If

'    Atan3 = Atan3          '   0 < Atan2 < 2π
'    Atan3 = Atan3 - pi     ' -π < Atan2 < π

End Function

Function arcSin(x As Double)       '-π/2< arcSin <π/2        -π/2< Atn <π/2
  If x <= -1 Then
    arcSin = 3 * pi / 2
  ElseIf x >= 1 Then
    arcSin = pi / 2
  Else
    '-π/2< arcSin <π/2
    arcSin = Atn(x / Sqr(1 - x ^ 2))
  End If

     arcSin = arcSin - pi / 2   '  -π/2 < arcSin <π/2
'     arcSin = arcSin            '      0 < arcSin <π

End Function

Function arcCos(x As Double)        ' 0 < arcCos <π            -π/2< Atn <π/2
  If x <= -1 Then
    arcCos = pi
  ElseIf x >= 1 Then
    arcCos = 0
  Else
    arcCos = Atn(-x / Sqr(-x ^ 2 + 1)) + pi / 2
  End If

     arcCos = arcCos            '      0 < arcCos <π
'     arcCos = arcCos - pi / 2   '  -π/2 < arcCos <π/2

End Function


'======================================================== 【M      】
'  結果を、配列に保存
'　　　　配列を、直接Cellへ格納
'　　　　※注）配列とCellの行と列は逆
'========================================================

Public Sub Data_Strage_to_array_check_curve()       ' Check curve

Dim I As Long, J As Long
Dim I1 As Long, J1 As Long
Dim tmp_int As Long


'【Cellに書き出し】

'-- Data Title 部
'
        J = 2    '配列の開始列
        I = 1     '配列の開始行

    dw_c = J + 20   '　配列の列数
     ReDim Data_Strage(dw_c, dw_n + 3)

                 tmp_int = 0
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(0, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(0, i)
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(1, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[142]  ' curve_yw(1, i)
                 tmp_int = 2
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(2, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(2, i)
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(3, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[146]  ' curve_yw(1, i)
                 tmp_int = 4
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(4, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(4, i)
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(5, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[150]  ' curve_yw(5, i)
                 tmp_int = 6
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(6, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(6, i)
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(7, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[154]   ' curve_yw(7, i)
                 tmp_int = 8
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int & ", j)"       ' curve_xw(8, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int & ", j)"       ' curve_yw(8, i)
        I = I + 1:            Data_Strage(I, J) = "xw(" & tmp_int + 1 & ", j)"   ' curve_xw(9, i)
        I = I + 1:            Data_Strage(I, J) = "yw(" & tmp_int + 1 & ", j)"   '[158]   ' curve_yw(9, i)


'--------------------------
'-- Data 部
'--------------------------
    For J = 0 To div_n      ' div Index 0 to 360

'------ 139 [Data]  Arc line of area that is the center of chamber area 】
    I = 1
                tmp_int = 0
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)                 ' xfo(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)                 ' yfo(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)          ' xfi(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)          ' yfi(j)     '[142]
                tmp_int = 2
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)                 ' xfo(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)                 ' yfo(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)          ' xfi(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)          ' yfi(j)     '[146]

                tmp_int = 4
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)                 ' xfo2(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)                 ' yfo2(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)          ' xfi2(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)          ' yfi2(j)     '[150]
                tmp_int = 6
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)              '  xmo2(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)              ' ymo2(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)      ' xmi2(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)      ' ymi2(j)     '[154]
                tmp_int = 8
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int, J)              ' xfo3(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int, J)              ' yfo3(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_xw(tmp_int + 1, J)      ' xfi3(j)
        I = I + 1:            Data_Strage(I, J + 3) = curve_yw(tmp_int + 1, J)      ' yfi3(j)     '[158]

    Next J


'-- Data 一括貼付

    DataSheetName = "DataSheet_6"

    Sheets(DataSheetName).Select
'    Sheets(DataSheetName).Cells.ClearContents                    '：Sheet全Cellの数式、文字列をクリア
'    Sheets(DataSheetName).Range("A1:FO999").ClearContents         '：指定Cellの数式、文字列をクリア

    I1 = 1                 ' 貼付先の先頭セルの、行と列 (i1, j1)
    J1 = 138
        With Sheets(DataSheetName)
            .Range(Cells(I1, J1), Cells(dw_n + 3 + I1, dw_c + J1)).Value _
                = WorksheetFunction.Transpose(Data_Strage)
        End With

        With Sheets(DataSheetName)
            .Cells(2, J1 + 2).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")
        End With

End Sub


'======================================================== 【      】
'  Get_Radius_Curvature
'     Call Get_Radius_curvature_min
'     use with Function :
'        Fn_Radius_Curvature(Rc_x As Double,  Rc_k As Double)
'
'        w_Rcurv_base(I, 0) = Fn_Radius_Curvature(phi_x, Rc_k)
'
'========================================================

Public Sub Get_Radius_Curvature()

    Dim I As Long, J As Long
    Dim Imax As Long, Jmax As Long
    Dim Phi_x As Double
    Dim istep As Double
    Dim phi_xmax As Double
    Dim Rc_a As Double, Rc_k As Double, Rc_g1 As Double, Rc_g2 As Double

    Dim Radius_Curvature_min(4) As Double

    Dim tmp_y As Double   ', tmp_ymin As Double
    Dim I1 As Long, J1 As Long

'------------------
'　　　：設定
'------------------

    phi_xmax = 1            ' [rad]            ' Divied angle
    istep = 0.002
    Imax = phi_xmax / istep
    Jmax = 1                ' column

    ReDim w_Rcurv_base(Jmax, Imax)
    ReDim w_Rcurv_base_min(Jmax)

'------------------
    Rc_a = a
    Rc_k = k
    Rc_g1 = g1
    Rc_g2 = g2

'------------------
'   g1,g2 固定値の場合、開始角度の値は、下記を読み込む。(計算短縮用)
'
  If g1 = 1.46 And g2 = 3.06 Then
           Wrap_Start_angle_min(1) = 0                 'FS_in   Wrap_Start_angle_min FS in (g1) Phi_1
           Wrap_Start_angle_min(2) = 0                 'FS_out  Wrap_Start_angle_min FS out (g1) Phi_2
           Wrap_Start_angle_min(3) = 0.116898192869584 'OS_in   Wrap_Start_angle_min OS in (g2) Phi_1
           Wrap_Start_angle_min(4) = 1.44514370435939  'OS_out  Wrap_Start_angle_min OS out (g2) Phi_2

           Radius_Curvature_min(1) = 1.63312393531954E+16 ' FS_in Radius_Curvature_min FS out (g1) Phi_2
           Radius_Curvature_min(2) = 1.63312393531954E+16 ' FS_out Radius_Curvature_min FS in (g1) Phi_1
           Radius_Curvature_min(3) = 0.53482025666059    ' OS_in Radius_Curvature_min OS in (g2) Phi_1
           Radius_Curvature_min(4) = 1.25922602486306    ' OS_out Radius_Curvature_min OS out (g2) Phi_2

  ElseIf g1 = 3.06 And g2 = 1.46 Then
           Wrap_Start_angle_min(1) = 0.116898192869584 'FS_in   Wrap_Start_angle_min FS in (g1) Phi_1
           Wrap_Start_angle_min(2) = 1.44514370435939  'FS_out  Wrap_Start_angle_min FS out (g1) Phi_2
           Wrap_Start_angle_min(3) = 0                 'OS_in   Wrap_Start_angle_min OS in (g2) Phi_1
           Wrap_Start_angle_min(4) = 0                 'OS_out  Wrap_Start_angle_min OS out (g2) Phi_2

           Radius_Curvature_min(1) = 0.53482025666059    ' FS_in Radius_Curvature_min FS out (g1) Phi_2
           Radius_Curvature_min(2) = 1.25922602486306    ' FS_out Radius_Curvature_min FS in (g1) Phi_1
           Radius_Curvature_min(3) = 1.63312393531954E+16   ' OS_in Radius_Curvature_min OS in (g2) Phi_1
           Radius_Curvature_min(4) = 1.63312393531954E+16    ' OS_out Radius_Curvature_min OS out (g2) Phi_2

  ElseIf g1 = 2.26 And g2 = 2.26 Then     ' g1 = g1 =2.26
        ' <Result>
           Wrap_Start_angle_min(1) = 1.00660817056929E-02 'FS_in   Wrap_Start_angle_min FS in (g1) Phi_1
           Wrap_Start_angle_min(2) = 0.710188517448186     'FS_out  Wrap_Start_angle_min FS out (g1) Phi_2
           Wrap_Start_angle_min(3) = 1.00660817056929E-02 'OS_in   Wrap_Start_angle_min OS in (g2) Phi_1
           Wrap_Start_angle_min(4) = 0.710188517448186    'OS_out  Wrap_Start_angle_min OS out (g2) Phi_2

           Radius_Curvature_min(1) = 0.655913190557202    ' FS_in Radius_Curvature_min FS out (g1) Phi_2
           Radius_Curvature_min(2) = 0.707351403897201    ' FS_out Radius_Curvature_min FS in (g1) Phi_1
           Radius_Curvature_min(3) = 0.655913190557202    ' OS_in Radius_Curvature_min OS in (g2) Phi_1
           Radius_Curvature_min(4) = 0.707351403897201    ' OS_out Radius_Curvature_min OS out (g2) Phi_2
  Else

    '--------
      For J = 1 To Jmax
          w_Rcurv_base_min(J) = 100           ' Initial value
      Next J
    '--------
      For I = 0 To Imax
         w_Rcurv_base(0, I) = I * istep      '[rad]  index No.
      Next I

    '------------------
    '     get w_Rcurv_base_min(J)
    '------------------
      For J = 1 To Jmax
          Rc_k = Rc_k + 0.1 * (J - 1)

          If Rc_k < 1 Then

          '--- calc minimum carvature of baseline
            Call Get_Radius_curvature_min                       ' Wrap_Start_angle_min(1)～(4)を決定

                w_Rcurv_base_min(J) = Rcurvature_min_b          ' result of minimum carvature
                Wrap_Start_angle_min_b = Wrap_Start_angle_min_b ' result of angle at minimum carvature

          ElseIf Rc_k = 1 Then
                      w_Rcurv_base_min(J) = 0.5
          Else
                      w_Rcurv_base_min(J) = 0
          End If
      Next J

  End If
          '--- set minimum carvature to Wrap_Start_angle_min(0)
            Wrap_Start_angle_min(0) = 1000
            For I = 1 To 4
                If Wrap_Start_angle_min(I) < Wrap_Start_angle_min(0) Then
                    Wrap_Start_angle_min(0) = Wrap_Start_angle_min(I)
                End If
            Next I

'---------------------
'   whole Wrap from the head start to the tail end
'---------------------

  Dim Phi_fi1() As Double
  Dim Phi_fo1() As Double
  Dim Phi_mi1() As Double
  Dim Phi_mo1() As Double

  Dim Iw_max(4) As Long
  Dim Aw_end(4) As Double

      Aw_end(1) = FS_in_end
      Aw_end(2) = FS_out_end
      Aw_end(3) = OS_in_end
      Aw_end(4) = OS_out_end

    ' 各Wrapの配列数を求める　　str,endが異なるため
      For I = 1 To 4
          Iw_max(I) = 0
          Iw_max(I) = Round((Aw_end(I) - Wrap_Start_angle_min(I)) / dw)

            If dw * Iw_max(I) < (Aw_end(I) - Wrap_Start_angle_min(I)) Then
               Iw_max(I) = Iw_max(I) + 1

            ElseIf dw * Iw_max(I) = (Aw_end(I) - Wrap_Start_angle_min(I)) Then
               Iw_max(I) = Iw_max(I)

            ElseIf dw * Iw_max(I) < (Aw_end(I) - Wrap_Start_angle_min(I)) Then
              Stop
               Iw_max(I) = Iw_max(I) - 1

            End If

      Next I

' For whole Wrap from the head start to the tail end

  ReDim xfi1(Iw_max(1)):   ReDim yfi1(Iw_max(1)):
  ReDim xfo1(Iw_max(2)):   ReDim yfo1(Iw_max(2)):
  ReDim xmi1(Iw_max(3)):   ReDim ymi1(Iw_max(3)):
  ReDim xmo1(Iw_max(4)):   ReDim ymo1(Iw_max(4)):

  ReDim Phi_fi1(Iw_max(1))
  ReDim Phi_fo1(Iw_max(2))
  ReDim Phi_mi1(Iw_max(3))
  ReDim Phi_mo1(Iw_max(4))

  '----Wrap FS_in　座標計算
          For I = 0 To Iw_max(1)
              Phi_x = FS_in_end - dw * I
                If Phi_x <= Wrap_Start_angle_min(1) Then
                   Phi_x = Wrap_Start_angle_min(1)
                  ' Stop
                  ' ElseIf I = Iw_max(1) Then
                  '  Phi_x = Wrap_Start_angle_min(1)
                End If

              Phi_fi1(I) = Phi_x
              xfi1(I) = Fn_xfi(Phi_x) + dx
              yfi1(I) = Fn_yfi(Phi_x) + dy

         Next I

  '----Wrap FS_out
         For I = 0 To Iw_max(2)
              Phi_x = FS_out_end - dw * I
                If Phi_x <= Wrap_Start_angle_min(2) Then
                   Phi_x = Wrap_Start_angle_min(2)
                  ' Stop
                  ' ElseIf I = Iw_max(2) Then
                  '  Phi_x = Wrap_Start_angle_min(2)
                End If

              Phi_fo1(I) = Phi_x
              xfo1(I) = Fn_xfo(Phi_x) + dx
              yfo1(I) = Fn_yfo(Phi_x) + dy
         Next I

  '----Wrap OS_in
         For I = 0 To Iw_max(3)
              Phi_x = OS_in_end - dw * I
                If Phi_x <= Wrap_Start_angle_min(3) Then
                   Phi_x = Wrap_Start_angle_min(3)
                  ' Stop
                  ' ElseIf I = Iw_max(3) Then
                  '   Phi_x = Wrap_Start_angle_min(3)
                End If

              Phi_mi1(I) = Phi_x
              xmi1(I) = Fn_xmi(Phi_x) + dx
              ymi1(I) = Fn_ymi(Phi_x) + dy
         Next I

  '----Wrap OS_out
         For I = 0 To Iw_max(4)
              Phi_x = OS_out_end - dw * I
                If Phi_x <= Wrap_Start_angle_min(4) Then
                   Phi_x = Wrap_Start_angle_min(4)
                  ' Stop
                  ' ElseIf I = Iw_max(4) Then
                  '   Phi_x = Wrap_Start_angle_min(4)
                End If

              Phi_mo1(I) = Phi_x
              xmo1(I) = Fn_xmo(Phi_x) + dx
              ymo1(I) = Fn_ymo(Phi_x) + dy
         Next I


'--------------------------
'-- Data 列貼付
'--------------------------

    DataSheetName = "DataSheet_6"

    Sheets(DataSheetName).Select
'    Sheets(DataSheetName).Range("HR4:HG1999").ClearContents     '：指定Cellの数式、文字列をクリア

    With Sheets(DataSheetName)

      I1 = 4                          'Range("A4").row            ' 貼付先の先頭セルの、行と列 (i1, j1)
      Imax = Iw_max(1)
      J1 = Range("HR1").Column        'Range("B1").Column
      Jmax = 0

        .Range(Cells(I1, J1), Cells(I1, J1 + 16)).ClearContents   '：指定Cellの数式、文字列をクリア
        .Cells(1, J1).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")
        .Cells(3, J1).Value = "Phi_fi"
        .Cells(3, J1 + 1).Value = "xfi"
        .Cells(3, J1 + 2).Value = "yfi"

        .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
            = WorksheetFunction.Transpose(Phi_fi1)
         .Range(Cells(I1, J1 + 1), Cells(Imax + I1, Jmax + J1 + 1)).Value _
            = WorksheetFunction.Transpose(xfi1)
         .Range(Cells(I1, J1 + 2), Cells(Imax + I1, Jmax + J1 + 2)).Value _
            = WorksheetFunction.Transpose(yfi1)


      Imax = Iw_max(2)
      J1 = J1 + 3
        .Cells(3, J1).Value = "Phi_fo"
        .Cells(3, J1 + 1).Value = "xfo"
        .Cells(3, J1 + 2).Value = "yfo"

        .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
            = WorksheetFunction.Transpose(Phi_fo1)
        .Range(Cells(I1, J1 + 1), Cells(Imax + I1, Jmax + J1 + 1)).Value _
            = WorksheetFunction.Transpose(xfo1)
         .Range(Cells(I1, J1 + 2), Cells(Imax + I1, Jmax + J1 + 2)).Value _
            = WorksheetFunction.Transpose(yfo1)

      Imax = Iw_max(3)
      J1 = J1 + 3
        .Cells(3, J1).Value = "Phi_mi"
        .Cells(3, J1 + 1).Value = "xmi"
        .Cells(3, J1 + 2).Value = "ymi"

        .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
            = WorksheetFunction.Transpose(Phi_mi1)
        .Range(Cells(I1, J1 + 1), Cells(Imax + I1, Jmax + J1 + 1)).Value _
            = WorksheetFunction.Transpose(xmi1)
         .Range(Cells(I1, J1 + 2), Cells(Imax + I1, Jmax + J1 + 2)).Value _
            = WorksheetFunction.Transpose(ymi1)

      Imax = Iw_max(4)
      J1 = J1 + 3
        .Cells(3, J1).Value = "Phi_mo"
        .Cells(3, J1 + 1).Value = "xmo"
        .Cells(3, J1 + 2).Value = "ymo"

        .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
            = WorksheetFunction.Transpose(Phi_mo1)
        .Range(Cells(I1, J1 + 1), Cells(Imax + I1, Jmax + J1 + 1)).Value _
            = WorksheetFunction.Transpose(xmo1)
         .Range(Cells(I1, J1 + 2), Cells(Imax + I1, Jmax + J1 + 2)).Value _
            = WorksheetFunction.Transpose(ymo1)

    End With

End Sub



'======================================================== 【      】
'  Get_Radius_curvature_min()
'
'     use with Function :
'        Fn_Radius_Curvature(Rc_x As Double, Rc_k As Double)
'
'========================================================

Public Sub Get_Radius_curvature_min()

    Dim I As Long         ' , J As Long
    Dim Imax As Long      ' , Jmax As Long
    Dim Phi_x As Double
    Dim istep As Double
    Dim phi_xmax As Double
    Dim Rc_x As Double, Rc_a As Double, Rc_k As Double, Rc_g1 As Double, Rc_g2 As Double

    Dim Radius_Curvature_min(4) As Double
    Dim tmp_C(4) As Double
'    Dim tmp_y As Double, tmp_ymin As Double, tmp_del As Double

'------------------
'
'------------------
    DataSheetName = "DataSheet_1"

    Sheets(DataSheetName).Activate

    phi_xmax = 1            ' [rad]            ' Divied angle
    istep = 0.002
    Imax = phi_xmax / istep
    Rc_k = k
    Rc_a = a

    Sheets(DataSheetName).Range("ZY1:AAB19").ClearContents     '：指定Cellの数式、文字列をクリア

      With Sheets(DataSheetName)
        .Range("ZY1").Value = k                ' Algebraic constat
        .Range("ZY2").Value = a                ' Algebraic constat
        .Range("ZY3").Value = g1               ' Algebraic constat
        .Range("ZY4").Value = g2               ' Algebraic constat
        .Range("ZY5").Value = qq               ' Algebraic constat

'        .Range("ZY8").Value = "Radius_Curvature_min"
'        .Range("ZY9").Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

       End With

'------------------
'  [1]   solver :   minimum Radius Carvature of base
'------------------
  If Rc_k > 1 Then
      Rcurvature_min_b = 0

  ElseIf Rc_k = 1 Then
      Rcurvature_min_b = 1 / 2

  ElseIf Rc_k < 0 Then
      Stop

  Else

      With Sheets(DataSheetName)
        .Range("ZZ1").Value = 0.5             ' Phi_X

        .Range("AAA2").Formula = "=(ZZ1 ^ (2 * ZY1) + ZY1 ^ 2 * ZZ1 ^ (2 * (ZY1 - 1))) ^ (1.5)"
        .Range("AAA3").Formula = "=(ZZ1 ^ (2 * ZY1) + 2 * ZY1 ^ 2 * ZZ1 ^ (2 * (ZY1 - 1)) - ZY1 * (ZY1 - 1) * ZZ1 ^ (2 * ZY1 - 2))"
        .Range("AAA1").Formula = "=(AAA2 / AAA3)"

        ' Fn_Rcurvature_baseline = (Rc_x ^ (2 * Rc_k) + Rc_k ^ 2 * Rc_x ^ (2 * (Rc_k - 1))) ^ (1.5) _
            / (Rc_x ^ (2 * Rc_k) + 2 * Rc_k ^ 2 * Rc_x ^ (2 * (Rc_k - 1)) - Rc_k * (Rc_k - 1) * Rc_x ^ (2 * Rc_k - 2))

      End With

        Sheets(DataSheetName).Select
        SolverReset

            '--------------------------
            '　制約条件  SolverAddで制約条件を設定。
            '     Relationは１が less than(≦)、２がequal＝、３が more than≧
            '--------------------------
            '        SolverOptions MaxTime:=0, Iterations:=200, Precision:=0.0000000000001, _
            '            Convergence:=0.00001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, _
            '            Derivatives:=2
            '        SolverAdd CellRef:="$AAA$2", Relation:=1, FormulaText:="12"
            '        SolverAdd CellRef:="$AAA$2", Relation:=3, FormulaText:="0.001"

        SolverOptions MaxTime:=20, Precision:=1E-16

      '　制約条件  SolverAddで制約条件を設定。
      '     Relationは１が less than(≦)、２がequal＝、３が more than≧
        SolverAdd CellRef:="ZZ1", Relation:=3, FormulaText:="0.00000000000001"
        SolverAdd CellRef:="ZZ1", Relation:=1, FormulaText:="PI()/2.5"

      ' Parameter設定
      '     MaxMinVal =2 （最小値にする=2；最大にする=1；特定値にする=3）
        SolverOk SetCell:="$AAA$1", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$ZZ$1"), _
            Engine:=1, EngineDesc:="GRG Nonlinear"

        SolverSolve UserFinish:=True    ' 結果ボックスを非表示

      '------------------
      '--- Results :   minimum Radius Carvature of base

        Wrap_Start_angle_min_b = Sheets(DataSheetName).Range("ZZ1").Value      ' result of minimum carvature
        Rcurvature_min_b = Sheets(DataSheetName).Range("AAA1").Value           ' result of angle at minimum carvature

      ' [1] Base Line : ---

       With Sheets(DataSheetName)
        .Range("AAD2").Value = "[1] Base Line : Radius_Curvature_min"
        .Range("AAD3").Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

        .Range("AAD4").Value = "k"                ' Algebraic constat
        .Range("AAD5").Value = "a"                ' Algebraic constat
        .Range("AAD6").Value = "g1"               ' Algebraic constat
        .Range("AAD7").Value = "g2"              ' Algebraic constat
        .Range("AAD8").Value = "Phi_min_b"                        ' result of minimum carvature
        .Range("AAD9").Value = "Rcurvature_min_b"                 ' result of angle at minimum carvaturet

        .Range("AAE4").Value = k                ' Algebraic constat
        .Range("AAE5").Value = a                ' Algebraic constat
        .Range("AAE6").Value = g1               ' Algebraic constat
        .Range("AAE7").Value = g2               ' Algebraic constat
        .Range("AAE8").Value = Wrap_Start_angle_min_b             ' result of minimum carvature
        .Range("AAE9").Value = Rcurvature_min_b                   ' result of angle at minimum carvaturet

       End With

  End If



'------------------
'  条件1) 外壁より小さい内壁側の曲率半径が、基本螺旋曲線の曲率半径より大きい場合
'　　　　　⇒ [2] 内壁側の曲率半径と等しくなる基本螺旋曲線の伸開角φi_minを求める
'          ⇒ [3] 内外壁の巻始め側の交点を求める。 条件1)の場合は、内外壁交点の伸開角は異なる
'------------------

'------------------
'  [2-1] FSout = Base_in : calc start angel at Radius_Curvature_min : Rcurvature_min_b=g1,g2
'------------------

   If Rc_k < 1 And Rcurvature_min_b < (g1 / a) Then          ' Wp_xfi

        Sheets(DataSheetName).Range("ZZ1:AAB19").ClearContents      '：指定Cellの数式、文字列をクリア

          With Sheets(DataSheetName)
            .Range("ZZ1").Value = Wrap_Start_angle_min_b            ' Phi_i : φ2in Initial value    Wp_xfi
            .Range("ZZ3").Value = Wrap_Start_angle_min_b

        .Range("AAA2").Formula = "=(ZZ1 ^ (2 * ZY1) + ZY1 ^ 2 * ZZ1 ^ (2 * (ZY1 - 1))) ^ (1.5)"
        .Range("AAA3").Formula = "=(ZZ1 ^ (2 * ZY1) + 2 * ZY1 ^ 2 * ZZ1 ^ (2 * (ZY1 - 1)) - ZY1 * (ZY1 - 1) * ZZ1 ^ (2 * ZY1 - 2))"
        .Range("AAA1").Formula = "=(AAA2 / AAA3)-(ZY3 / ZY2)"

          End With

        Sheets(DataSheetName).Select
        SolverReset

            '--------------------------
            '　制約条件  SolverAddで制約条件を設定。
            '     Relation : １が less than(≦)、２がequal＝、３が more than≧

           SolverOptions MaxTime:=20, Precision:=1E-16

           SolverAdd CellRef:="ZZ1", Relation:=3, FormulaText:="ZZ3"
           SolverAdd CellRef:="ZZ1", Relation:=1, FormulaText:="PI()*0.99"

           SolverOk SetCell:="$AAA$1", MaxMinVal:=3, ValueOf:=0, ByChange:=Range("$ZZ$1"), Engine _
               :=1, EngineDesc:="GRG Nonlinear"
           SolverSolve UserFinish:=True

        Wrap_Start_angle_min_g1 = Sheets(DataSheetName).Range("ZZ1").Value      '
        Rcurvature_min_g1 = Sheets(DataSheetName).Range("AAA1").Value - g1 / a  '

    Else
        Wrap_Start_angle_min_g1 = 0
            Rc_x = Wrap_Start_angle_min_g1
        Rcurvature_min_g1 = Fn_Radius_curvature(Rc_x, Rc_k) - g1 / a

   End If

   ' 結果表示

       With Sheets(DataSheetName)
        .Range("AAD12").Value = "[2-1] FSout = Base_in : Radius_Curvature_min"
        .Range("AAD13").Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

        .Range("AAD14").Value = "k"                ' Algebraic constat
        .Range("AAD15").Value = "a"                ' Algebraic constat
        .Range("AAD16").Value = "g1"               ' Algebraic constat
        .Range("AAD17").Value = "g2"              ' Algebraic constat
        .Range("AAD18").Value = "Phi_min_g1"                      ' Wrap_Start_angle_min FS out (g1)
        .Range("AAD19").Value = "Rcurvature_min_g1"               ' Radius_Curvature_min FS out (g1)

        .Range("AAE14").Value = k                ' Algebraic constat
        .Range("AAE15").Value = a                ' Algebraic constat
        .Range("AAE16").Value = g1               ' Algebraic constat
        .Range("AAE17").Value = g2               ' Algebraic constat
        .Range("AAE18").Value = Wrap_Start_angle_min_g1           ' Wrap_Start_angle_min FS out (g1)
        .Range("AAE19").Value = Rcurvature_min_g1                 ' Radius_Curvature_min FS out (g1)

       End With


'------------------
'  [2-2] OSout = Base_in : calc start angel at Radius_Curvature_min : Rcurvature_min_b=g1,g2
'------------------

   If Rc_k < 1 And Rcurvature_min_b < (g2 / a) Then          ' Wp_xmi

        Sheets(DataSheetName).Range("ZZ1:AAB19").ClearContents      '：指定Cellの数式、文字列をクリア

          With Sheets(DataSheetName)
            .Range("ZZ1").Value = Wrap_Start_angle_min_b          ' Phi_i : for guzai2 Initial value    ' Wp_xmi
            .Range("ZZ3").Value = Wrap_Start_angle_min_b

        .Range("AAA2").Formula = "=(ZZ1 ^ (2 * ZY1) + ZY1 ^ 2 * ZZ1 ^ (2 * (ZY1 - 1))) ^ (1.5)"
        .Range("AAA3").Formula = "=(ZZ1 ^ (2 * ZY1) + 2 * ZY1 ^ 2 * ZZ1 ^ (2 * (ZY1 - 1)) - ZY1 * (ZY1 - 1) * ZZ1 ^ (2 * ZY1 - 2))"
        .Range("AAA1").Formula = "=(AAA2 / AAA3)-(ZY4 / ZY2)"

          End With

        Sheets(DataSheetName).Select
        SolverReset

            '--------------------------
            '　制約条件  SolverAddで制約条件を設定。
            '     Relationは１が less than(≦)、２がequal＝、３が more than≧
            '--------------------------

        SolverOptions MaxTime:=20, Precision:=1E-16

        SolverAdd CellRef:="ZZ1", Relation:=3, FormulaText:="ZZ3"
        SolverAdd CellRef:="ZZ1", Relation:=1, FormulaText:="PI()*0.99"

        SolverOk SetCell:="$AAA$1", MaxMinVal:=3, ValueOf:=0, ByChange:=Range("$ZZ$1"), Engine _
            :=1, EngineDesc:="GRG Nonlinear"
        SolverSolve UserFinish:=True

        Wrap_Start_angle_min_g2 = Sheets(DataSheetName).Range("ZZ1").Value      '
        Rcurvature_min_g2 = Sheets(DataSheetName).Range("AAA1").Value - g2 / a          '

    Else
        Wrap_Start_angle_min_g2 = 0
            Rc_x = Wrap_Start_angle_min_g2
        Rcurvature_min_g2 = Fn_Radius_curvature(Rc_x, Rc_k) - g2 / a

   End If

   ' 結果表示

       With Sheets(DataSheetName)
        .Range("AAD22").Value = "[2-2] OSout = Base_in : Radius_Curvature_min"
        .Range("AAD23").Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

        .Range("AAD24").Value = "k"                ' Algebraic constat
        .Range("AAD25").Value = "a"                ' Algebraic constat
        .Range("AAD26").Value = "g1"               ' Algebraic constat
        .Range("AAD27").Value = "g2"              ' Algebraic constat
        .Range("AAD28").Value = "Phi_min_g2"                      ' Wrap_Start_angle_min OS out (g1)
        .Range("AAD29").Value = "Rcurvature_min_g2"               ' Radius_Curvature_min OS out (g1)

        .Range("AAE24").Value = k                ' Algebraic constat
        .Range("AAE25").Value = a                ' Algebraic constat
        .Range("AAE26").Value = g1               ' Algebraic constat
        .Range("AAE27").Value = g2               ' Algebraic constat
        .Range("AAE28").Value = Wrap_Start_angle_min_g2           ' Wrap_Start_angle_min OS out (g1)
        .Range("AAE29").Value = Rcurvature_min_g2                 ' Radius_Curvature_min OS out (g1)

       End With


'--------------------------
'　[3] 偏角φi、φoの収束解をSolverで求める
'     FS,OS内外壁の巻始め交点の(φi、φo)を求める
'　　 [1]の基本螺旋曲線の最小曲率半径＜g1,g2の場合、(φi、φo)が異なる
'--------------------------

'-----------------
'  [3-1] FS  Solver：FS  φi、φo
'-----------------

    If Rcurvature_min_b < (g1 / a) Then          ' Wp_xfi : for FS side φi、φo
      ' If Rc_k < 1 And Rcurvature_min_b < (g1 / a) Then          ' Wp_xfi : for FS side φi、φo

        Sheets(DataSheetName).Range("ZZ1:AAB19").ClearContents     '：指定Cellの数式、文字列をクリア

          With Sheets(DataSheetName)

            .Range("ZZ1").Value = Wrap_Start_angle_min_g1 * 2       '0.5       ' Phi_i : φ2in Initial value
            .Range("ZZ2").Value = Wrap_Start_angle_min_g1 * 2       '0.5       ' Phi_o : φ2out Initial value
            .Range("ZZ3").Value = Wrap_Start_angle_min_g1

            .Range("AAA2").Formula = "=ZY2 * ZZ1 ^ ZY1 * Cos(ZZ1 - ZY5) + ZY3 * Cos(ZZ1 - ZY5 - ATAN(ZY1 / ZZ1))"      ' Wp_xfi
            .Range("AAA3").Formula = "=ZY2 * ZZ1 ^ ZY1 * Sin(ZZ1 - ZY5) + ZY3 * Sin(ZZ1 - ZY5 - ATAN(ZY1 / ZZ1))"      ' Wp_yfi
            .Range("AAA4").Formula = "=-ZY2 * ZZ2 ^ ZY1 * Cos(ZZ2 - ZY5) + ZY3 * Cos(ZZ2 - ZY5 - ATAN(ZY1 / ZZ2))"     ' Wp_xfo
            .Range("AAA5").Formula = "=-ZY2 * ZZ2 ^ ZY1 * Sin(ZZ2 - ZY5) + ZY3 * Sin(ZZ2 - ZY5 - ATAN(ZY1 / ZZ2))"     ' Wp_yfo
            .Range("AAA1").Formula = "=(AAA2-AAA4)^2+(AAA3-AAA5)^2"
          End With

        Sheets(DataSheetName).Select
        SolverReset

      '--------------------------
      '　制約条件  SolverAddで制約条件を設定。
      '     Relationは１が less than(≦)、２がequal＝、３が more than≧
      '--------------------------
      '        SolverOptions MaxTime:=0, Iterations:=200, Precision:=0.0000000000001, _
      '            Convergence:=0.00001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, _
      '            Derivatives:=2
      '        SolverAdd CellRef:="$AAA$2", Relation:=1, FormulaText:="12"
      '        SolverAdd CellRef:="$AAA$2", Relation:=3, FormulaText:="0.001"

        SolverOptions MaxTime:=20, Precision:=1E-16
'        SolverOptions MaxTime:=20, Precision:=1E-16

        SolverAdd CellRef:="ZZ1", Relation:=3, FormulaText:="0.000001"                  ' Phi_1 xfi Lower limit
        SolverAdd CellRef:="ZZ1", Relation:=1, FormulaText:="PI()*0.99"                  ' Phi_1 xfi Upper limit

        SolverAdd CellRef:="ZZ2", Relation:=3, FormulaText:="ZZ3"           ' Phi_2 xfo Lower limit
        SolverAdd CellRef:="ZZ2", Relation:=1, FormulaText:="PI()*0.99"                  ' Phi_2 xfo Upper limit


        SolverOk SetCell:="$AAA$1", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$ZZ$1", "$ZZ$2"), Engine _
            :=1, EngineDesc:="GRG Nonlinear"
        SolverSolve UserFinish:=True

    '-----------------
    '   結果：FS  φi、φo
    '-----------------

        Wrap_Start_angle_min(1) = Sheets(DataSheetName).Range("ZZ1").Value       ' Wrap_Start_angle_min FS in (g1) Phi_1
                Rc_x = Wrap_Start_angle_min(1)
        Radius_Curvature_min(1) = Fn_Radius_curvature(Rc_x, Rc_k)                ' Radius_Curvature_min FS in (g1) Phi_1

        Wrap_Start_angle_min(2) = Sheets(DataSheetName).Range("ZZ2").Value       ' Wrap_Start_angle_min FS out (g1) Phi_2
                Rc_x = Wrap_Start_angle_min(2)
        Radius_Curvature_min(2) = Fn_Radius_curvature(Rc_x, Rc_k)                ' Wrap_Start_angle_min FS out (g1) Phi_2

            '  Radius_Curvature_min(3) = w_Rcurv_base_min(J) * Rc_a + g2 ' OS in
            '  Radius_Curvature_min(4) = w_Rcurv_base_min(J) * Rc_a - g2 ' OS out
            '  Wrap_Start_angle_min(3) = arrReturn(0)  'OS_in
            '  Wrap_Start_angle_min(4) = arrReturn(1)  'OS_out

    Else
'        Stop
        Wrap_Start_angle_min(1) = 0                                         ' Wrap_Start_angle_min FS in (g1) Phi_1
        Radius_Curvature_min(1) = Fn_Radius_curvature(0, Rc_k)              ' Radius_Curvature_min FS in (g1) Phi_1

        Wrap_Start_angle_min(2) = 0                                         ' Wrap_Start_angle_min FS out (g1) Phi_2
        Radius_Curvature_min(2) = Fn_Radius_curvature(0, Rc_k)              ' Radius_Curvature_min FS out (g1) Phi_2

    End If

   ' 結果表示　[3-FS] --- φi、φo

       With Sheets(DataSheetName)
        .Range("AAD32").Value = "[3-1] FS Radius_Curvature_min"
        .Range("AAD33").Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

        .Range("AAD34").Value = "k"                ' Algebraic constat
        .Range("AAD35").Value = "a"                ' Algebraic constat
        .Range("AAD36").Value = "g1"               ' Algebraic constat
        .Range("AAD37").Value = "g2"              ' Algebraic constat
        .Range("AAD38").Value = "Wrap_Start_angle_min(1) FSin = Base_out"       ' Wrap_Start_angle_min FS in (g1) Phi_1
        .Range("AAD39").Value = "Radius_Curvature_min(1) FSin = Base_out"       ' Radius_Curvature_min FS in (g1) Phi_1
        .Range("AAD40").Value = "Wrap_Start_angle_min(2) FSout = Base_in"       ' Wrap_Start_angle_min FS out (g1) Phi_2
        .Range("AAD41").Value = "Radius_Curvature_min(2) FSout = Base_in"       ' Radius_Curvature_min FS out (g1) Phi_2

        .Range("AAE34").Value = k                ' Algebraic constat
        .Range("AAE35").Value = a                ' Algebraic constat
        .Range("AAE36").Value = g1               ' Algebraic constat
        .Range("AAE37").Value = g2               ' Algebraic constat
        .Range("AAE38").Value = Wrap_Start_angle_min(1)               ' Wrap_Start_angle_min FS in (g1) Phi_1
        .Range("AAE39").Value = Radius_Curvature_min(1)               ' Radius_Curvature_min FS in (g1) Phi_1
        .Range("AAE40").Value = Wrap_Start_angle_min(2)               ' Wrap_Start_angle_min FS out (g1) Phi_2
        .Range("AAE41").Value = Radius_Curvature_min(2)               ' Radius_Curvature_min FS out (g1) Phi_2

       End With


'-----------------
' [3-2] OS Solver：OS  φi、φo
'-----------------

    If Rcurvature_min_b < (g2 / a) Then          ' Wp_xmi    ' solve for OS side φi、φo
      ' If Rc_k < 1 And Rcurvature_min_b < g2 / a Then          ' Wp_xmi    ' solve for OS side φi、φo

        Sheets(DataSheetName).Range("ZZ1:AAB19").ClearContents     '：指定Cellの数式、文字列をクリア

          With Sheets(DataSheetName)

            .Range("ZZ1").Value = Wrap_Start_angle_min_g2 * 2           ' Phi_i : φ2in Initial value
            .Range("ZZ2").Value = Wrap_Start_angle_min_g2 * 2           ' Phi_o : φ2out Initial value
            .Range("ZZ3").Value = Wrap_Start_angle_min_g2

            .Range("AAA12").Formula = "=-ZY2 * ZZ1 ^ ZY1 * Cos(ZZ1 - ZY5) - ZY4 * Cos(ZZ1 - ZY5 - ATAN(ZY1 / ZZ1))"     ' Wp_xmi
            .Range("AAA13").Formula = "=-ZY2 * ZZ1 ^ ZY1 * Sin(ZZ1 - ZY5) - ZY4 * Sin(ZZ1 - ZY5 - ATAN(ZY1 / ZZ1))"     ' Wp_ymi
            .Range("AAA14").Formula = "=ZY2 * ZZ2 ^ ZY1 * Cos(ZZ2 - ZY5) - ZY4 * Cos(ZZ2 - ZY5 - ATAN(ZY1 / ZZ2))"     ' Wp_xmo
            .Range("AAA15").Formula = "=ZY2 * ZZ2 ^ ZY1 * Sin(ZZ2 - ZY5) - ZY4 * Sin(ZZ2 - ZY5 - ATAN(ZY1 / ZZ2))"     ' Wp_ymo
            .Range("AAA11").Formula = "=(AAA12-AAA14)^2+(AAA13-AAA15)^2"
          End With

        Sheets(DataSheetName).Select
        SolverReset

      '--------------------------
      '　制約条件  SolverAddで制約条件を設定。
      '     Relationは１が less than(≦)、２がequal＝、３が more than≧
      '--------------------------

        SolverOptions MaxTime:=20, Precision:=1E-16

        SolverAdd CellRef:="ZZ1", Relation:=3, FormulaText:="0.0000001"                ' Phi_1 xmi Lower limit
        SolverAdd CellRef:="ZZ1", Relation:=1, FormulaText:="PI()*0.99"                ' Phi_1 xmi Upper limit

        SolverAdd CellRef:="ZZ2", Relation:=3, FormulaText:="ZZ3"           ' Phi_2 xmo Lower limit
        SolverAdd CellRef:="ZZ2", Relation:=1, FormulaText:="PI()*0.99"                ' Phi_2 xmo Upper limit


        SolverOk SetCell:="$AAA$11", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$ZZ$1", "$ZZ$2"), Engine _
            :=1, EngineDesc:="GRG Nonlinear"
        SolverSolve UserFinish:=True

      '-----------------
      '   結果
      '-----------------

        Wrap_Start_angle_min(3) = Sheets(DataSheetName).Range("ZZ1").Value       ' Wrap_Start_angle_min OS in (g2) Phi_1
                Rc_x = Wrap_Start_angle_min(3)
        Radius_Curvature_min(3) = Fn_Radius_curvature(Rc_x, Rc_k)                ' Radius_Curvature_min OS in (g2) Phi_1

        Wrap_Start_angle_min(4) = Sheets(DataSheetName).Range("ZZ2").Value       ' Wrap_Start_angle_min OS out (g2) Phi_2
                Rc_x = Wrap_Start_angle_min(4)
        Radius_Curvature_min(4) = Fn_Radius_curvature(Rc_x, Rc_k)                ' Radius_Curvature_min OS out (g2) Phi_2

            '  Radius_Curvature_min(3) = w_Rcurv_base_min(J) * Rc_a + g2 ' OS in
            '  Radius_Curvature_min(4) = w_Rcurv_base_min(J) * Rc_a - g2 ' OS out
            '  Wrap_Start_angle_min(3) = arrReturn(0)  'OS_in
            '  Wrap_Start_angle_min(4) = arrReturn(1)  'OS_out

    Else
'        Stop
        Wrap_Start_angle_min(3) = 0                                         ' Wrap_Start_angle_min OS in (g2) Phi_1
        Radius_Curvature_min(3) = Fn_Radius_curvature(0, Rc_k)              ' Radius_Curvature_min OS in (g2) Phi_1

        Wrap_Start_angle_min(4) = 0                                         ' Wrap_Start_angle_min OS out (g2) Phi_2
        Radius_Curvature_min(4) = Fn_Radius_curvature(0, Rc_k)              ' Radius_Curvature_min OS out (g2) Phi_2


    End If

      '----------------
      '   結果表示
      '----------------

       With Sheets(DataSheetName)
        .Range("AAD42").Value = "[3-2] OS Radius_Curvature_min"
        .Range("AAD43").Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

        .Range("AAD44").Value = "k"                ' Algebraic constat
        .Range("AAD45").Value = "a"                ' Algebraic constat
        .Range("AAD46").Value = "g1"               ' Algebraic constat
        .Range("AAD47").Value = "g2"               ' Algebraic constat
        .Range("AAD48").Value = "Wrap_Start_angle_min(3) OSin = Base_out"       ' Wrap_Start_angle_min OS in (g2) Phi_1
        .Range("AAD49").Value = "Radius_Curvature_min(3) OSin = Base_out"       ' Radius_Curvature_min OS in (g2) Phi_1
        .Range("AAD50").Value = "Wrap_Start_angle_min(4) OSout = Base_in"       ' Wrap_Start_angle_min OS out (g2) Phi_2
        .Range("AAD51").Value = "Radius_Curvature_min(4) OSout = Base_in"       ' Radius_Curvature_min OS out (g2) Phi_2

        .Range("AAE44").Value = k                ' Algebraic constat
        .Range("AAE45").Value = a                ' Algebraic constat
        .Range("AAE46").Value = g1               ' Algebraic constat
        .Range("AAE47").Value = g2               ' Algebraic constat
        .Range("AAE48").Value = Wrap_Start_angle_min(3)               ' Wrap_Start_angle_min OS in (g2) Phi_1
        .Range("AAE49").Value = Radius_Curvature_min(3)               ' Radius_Curvature_min OS in (g2) Phi_1
        .Range("AAE50").Value = Wrap_Start_angle_min(4)               ' Wrap_Start_angle_min OS out (g2) Phi_2
        .Range("AAE51").Value = Radius_Curvature_min(4)               ' Radius_Curvature_min OS out (g2) Phi_2

      End With

End Sub



'========================================================
'  Fn_Radius_curvature(Rc_x As Double, Rc_k As Double) As Double
'
'========================================================

Public Function Fn_Radius_curvature(Rc_x As Double, Rc_k As Double) As Double

  If Rc_x = 0 And Rc_k < 1 Then
      Fn_Radius_curvature = Tan(pi / 2)

    ElseIf Rc_x = 0 And Rc_k = 1 Then
      Fn_Radius_curvature = 1 / 2

    ElseIf Rc_x = 0 And Rc_k > 1 Then
      Fn_Radius_curvature = 0

    Else
      Fn_Radius_curvature = (Rc_x ^ (2 * Rc_k) + Rc_k ^ 2 * Rc_x ^ (2 * (Rc_k - 1))) ^ (1.5) _
            / (Rc_x ^ (2 * Rc_k) + 2 * Rc_k ^ 2 * Rc_x ^ (2 * (Rc_k - 1)) - Rc_k * (Rc_k - 1) * Rc_x ^ (2 * Rc_k - 2))
  End If

End Function




'========================================================
'   Fn_Wrap_Start_angle_min  with Solver
'
'========================================================

Public Function Fn_Wrap_Start_angle_min(ByVal g1 As Double, ByVal g2 As Double, ByVal DataSheetName As String) As Variant

    Dim I As Long, J As Long
    Dim tmp_C(4) As Double

    DataSheetName = "DataSheet_1"

    Sheets(DataSheetName).Activate

'-----------------
' Solver   xs_in=xs_out, ys_in=ys_out  -> φs_in and φs_out
'   Solver の式と制約条件を設定する。
'   SolverReset
'   SolverOptions
'   SolverOk:       目標条件を設定する｡
'   SolverSolve:       ソルバーを実行する｡
'   SolverFinish: 終了処理。求めた解を該当セル(B1欄)に書き込む。
'-----------------

    '--------------------------
    '  set initioal values and constraint conditions
    '--------------------------
'        phi_i = 0.8   ' 初期値 =:= Pi/4
'        Phi_o = 0.8   ' 初期値 =:= Pi/4

      With Sheets(DataSheetName)

        .Range("ZY1").Value = k                ' Algebraic constat
        .Range("ZY2").Value = a                ' Algebraic constat
        .Range("ZY3").Value = g1               ' Algebraic constat
        .Range("ZY4").Value = g2               ' Algebraic constat
        .Range("ZY5").Value = qq               ' Algebraic constat

       End With

            '  Wp_xfi = a * phi_i ^ k * Cos(phi_i - qq) + g1 * Cos(phi_i - qq - Atn(k / phi_i)) + dx
            '  Wp_yfi = a * phi_i ^ k * Sin(phi_i - qq) + g1 * Sin(phi_i - qq - Atn(k / phi_i)) + dy
            '  Wp_xfo = -a * Phi_o ^ k * Cos(Phi_o - qq) + g1 * Cos(Phi_o - qq - Atn(k / Phi_o)) + dx
            '  Wp_yfo = -a * Phi_o ^ k * Sin(Phi_o - qq) + g1 * Sin(Phi_o - qq - Atn(k / Phi_o)) + dy
            '
            '  Wp_xmi = -a * Phi_o ^ k * Cos(Phi_o - qq) - g2 * Cos(Phi_o - qq - Atn(k / Phi_o)) + dx
            '  Wp_ymi = -a * Phi_o ^ k * Sin(Phi_o - qq) - g2 * Sin(Phi_o - qq - Atn(k / Phi_o)) + dy
            '  Wp_xmo = a * phi_i ^ k * Cos(phi_i - qq) - g2 * Cos(phi_i - qq - Atn(k / phi_i)) + dx
            '  Wp_ymo = a * phi_i ^ k * Sin(phi_i - qq) - g2 * Sin(phi_i - qq - Atn(k / phi_i)) + dy

          '  .Range("AAA2").Formula = "=(if zz1=1, ZY2 * ZZ1 ^ ZY1 * Cos(ZZ1 - ZY4) + ZY3 * Cos(ZZ1 - ZY4 - ATAN(ZY1 / ZZ1)), _
          '     ZY2 * ZZ1 ^ ZY1 * Cos(ZZ1 - ZY4) + ZY3 * Cos(ZZ1 - ZY4 - ATAN(ZY1 / ZZ1)))"     ' Wp_xfi


    '--------------------------
    '　偏角φi、φoの収束解をSolverで求める
    '--------------------------

    If g2 = 0 Then       ' for FS side φi、φo

          With Sheets(DataSheetName)
            .Range("ZZ1").Value = 0.5             ' Phi_i : φ2in Initial value
            .Range("ZZ2").Value = 0.5            ' Phi_o : φ2out Initial value
            .Range("ZZ3").Value = Wrap_Start_angle_min_g1

            .Range("AAA2").Formula = "=ZY2 * ZZ1 ^ ZY1 * Cos(ZZ1 - ZY5) + ZY3 * Cos(ZZ1 - ZY5 - ATAN(ZY1 / ZZ1))"     ' Wp_xfi
            .Range("AAA3").Formula = "=ZY2 * ZZ1 ^ ZY1 * Sin(ZZ1 - ZY5) + ZY3 * Sin(ZZ1 - ZY5 - ATAN(ZY1 / ZZ1))"     ' Wp_yfi
            .Range("AAA4").Formula = "=-ZY2 * ZZ2 ^ ZY1 * Cos(ZZ2 - ZY5) + ZY3 * Cos(ZZ2 - ZY5 - ATAN(ZY1 / ZZ2))"     ' Wp_xfo
            .Range("AAA5").Formula = "=-ZY2 * ZZ2 ^ ZY1 * Sin(ZZ2 - ZY5) + ZY3 * Sin(ZZ2 - ZY5 - ATAN(ZY1 / ZZ2))"     ' Wp_yfo
            .Range("AAA1").Formula = "=(AAA2-AAA4)^2+(AAA3-AAA5)^2"
          End With


        Sheets(DataSheetName).Select
        SolverReset

      '--------------------------
      '　制約条件  SolverAddで制約条件を設定。
      '     Relationは１が less than(≦)、２がequal＝、３が more than≧
      '--------------------------
      '        SolverOptions MaxTime:=0, Iterations:=200, Precision:=0.0000000000001, _
      '            Convergence:=0.00001, StepThru:=False, Scaling:=False, AssumeNonNeg:=False, _
      '            Derivatives:=2
      '        SolverAdd CellRef:="$AAA$2", Relation:=1, FormulaText:="12"
      '        SolverAdd CellRef:="$AAA$2", Relation:=3, FormulaText:="0.001"

        SolverOptions MaxTime:=20, Precision:=0.0000000000001

        SolverAdd CellRef:="ZZ1", Relation:=3, FormulaText:="0.00000000000001"
        SolverAdd CellRef:="ZZ1", Relation:=1, FormulaText:="PI()/2.5"

        SolverAdd CellRef:="ZZ2", Relation:=3, FormulaText:="ZZ3"   ' "0.00000000000001"
        SolverAdd CellRef:="ZZ2", Relation:=1, FormulaText:="PI()/2.5"


        SolverOk SetCell:="$AAA$1", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$ZZ$1", "$ZZ$2"), Engine _
            :=1, EngineDesc:="GRG Nonlinear"
        SolverSolve UserFinish:=True

        tmp_C(3) = Sheets(DataSheetName).Range("AAA1").Value      ' x収束誤差
        tmp_C(4) = Sheets(DataSheetName).Range("AAA2").Value      ' y収束誤差


    ElseIf g1 = 0 Then    ' solve for OS side φi、φo

          With Sheets(DataSheetName)
            .Range("ZZ1").Value = 0.5            ' Phi_i : φ2in Initial value
            .Range("ZZ2").Value = 0.5             ' Phi_o : φ2out Initial value
            .Range("ZZ3").Value = Wrap_Start_angle_min_g2

            .Range("AAA12").Formula = "=-ZY2 * ZZ1 ^ ZY1 * Cos(ZZ1 - ZY5) - ZY4 * Cos(ZZ1 - ZY5 - ATAN(ZY1 / ZZ1))"     ' Wp_xmi
            .Range("AAA13").Formula = "=-ZY2 * ZZ1 ^ ZY1 * Sin(ZZ1 - ZY5) - ZY4 * Sin(ZZ1 - ZY5 - ATAN(ZY1 / ZZ1))"     ' Wp_ymi
            .Range("AAA14").Formula = "=ZY2 * ZZ2 ^ ZY1 * Cos(ZZ2 - ZY5) - ZY4 * Cos(ZZ2 - ZY5 - ATAN(ZY1 / ZZ2))"     ' Wp_xmo
            .Range("AAA15").Formula = "=ZY2 * ZZ2 ^ ZY1 * Sin(ZZ2 - ZY5) - ZY4 * Sin(ZZ2 - ZY5 - ATAN(ZY1 / ZZ2))"     ' Wp_ymo
            .Range("AAA11").Formula = "=(AAA12-AAA14)^2+(AAA13-AAA15)^2"
          End With

        Sheets(DataSheetName).Select
        SolverReset

      '--------------------------
      '　制約条件  SolverAddで制約条件を設定。
      '     Relationは１が less than(≦)、２がequal＝、３が more than≧
      '--------------------------

        SolverOptions MaxTime:=20, Precision:=0.0000000000001

        SolverAdd CellRef:="ZZ1", Relation:=3, FormulaText:="0.00000000001"
        SolverAdd CellRef:="ZZ1", Relation:=1, FormulaText:="PI()/2.5"

        SolverAdd CellRef:="ZZ2", Relation:=3, FormulaText:="ZZ3"   '"0.00000000001"
        SolverAdd CellRef:="ZZ2", Relation:=1, FormulaText:="PI()/2.5"


        SolverOk SetCell:="$AAA$11", MaxMinVal:=2, ValueOf:=0, ByChange:=Range("$ZZ$1", "$ZZ$2"), Engine _
            :=1, EngineDesc:="GRG Nonlinear"
        SolverSolve UserFinish:=True

        tmp_C(3) = Sheets(DataSheetName).Range("AAA11").Value      ' x収束誤差
        tmp_C(4) = Sheets(DataSheetName).Range("AAA12").Value      ' y収束誤差

    Else
        Stop
    End If

        tmp_C(0) = Sheets(DataSheetName).Range("ZZ1").Value       ' phi_i   Wp_xfi
        tmp_C(1) = Sheets(DataSheetName).Range("ZZ2").Value       ' Phi_o   Wp_xfo
        tmp_C(2) = (tmp_C(0) - tmp_C(1)) / pi * 180               ' 差 [deg]

    '-----------------
    '   結果：φi、φo
    '       Fn_Wrap_Start_angle_min(0) = phi_i
    '       Fn_Wrap_Start_angle_min(1) = phi_0
    '-----------------

      Fn_Wrap_Start_angle_min = Array(tmp_C(0), tmp_C(1))
'          Fn_Wrap_Start_angle_min = Array(Range("ZZ1").Value, Range("ZZ2").Value)


End Function


'======================================================== 【      】
'  Get_Radius_Curvature_parameter
'
'     use with Function :
'        Fn_Radius_Curvature(Rc_x As Double, Rc_k As Double)
'
'        w_Rcurv_base(I, 0) = Fn_Radius_Curvature(phi_x, Rc_k)
'
'========================================================

Public Sub Get_Radius_Curvature_parameter()

    Dim I As Long, J As Long
    Dim Imax As Long, Jmax As Long
    Dim Phi_x As Double
    Dim istep As Double
    Dim phi_xmax As Double
    Dim Rc_a As Double, Rc_k As Double    ' , Rc_g As Double

'    Dim tmp_y As Double, tmp_ymin As Double
    Dim I1 As Long, J1 As Long

'------------------
'　　　：設定
'------------------

    phi_xmax = 1            ' [rad]            ' Divied angle
    istep = 0.002
    Imax = phi_xmax / istep
    Jmax = 6

    ReDim w_Rcurv_base(Jmax, Imax)
    ReDim w_Rcurv_base_min(Jmax)

'------------------
    Rc_a = a
    Rc_k = k
'    Rc_g = 0

'------------------------
'-- set Initial value
'------------------------

    For J = 1 To Jmax
        w_Rcurv_base_min(J) = 100                  ' Initial value
    Next J

    For I = 0 To Imax
       w_Rcurv_base(0, I) = I * istep      '[rad]
    Next I

    For J = 1 To Jmax
        Rc_k = 0.71 + 0.1 * (J - 1)

        For I = 0 To Imax
           Phi_x = I * istep
           w_Rcurv_base(J, I) = Fn_Radius_curvature(Phi_x, Rc_k)

          If Rc_k < 1 Then

              If w_Rcurv_base_min(J) > w_Rcurv_base(J, I) Then
                     w_Rcurv_base_min(J) = w_Rcurv_base(J, I)
              End If

          ElseIf Rc_k = 1 Then
                     w_Rcurv_base_min(J) = 1 / 2
          Else
                     w_Rcurv_base_min(J) = 0
          End If

        Next I

    Next J



'-- Data 一括貼付

    DataSheetName = "DataSheet_1"

    Sheets(DataSheetName).Select
'    Sheets(DataSheetName).Cells.Clear                         '：Sheet全Cellの値と書式を消去

'    Sheets(DataSheetName).Cells.ClearContents                 '：Sheet全Cellの値のみ消去 , ClearComments
'    Sheets(DataSheetName).Cells.ClearFormats                  '：Sheet全Cellの書式のみ消去, ClearOutline
    Sheets(DataSheetName).Range("A1:KO999").ClearContents     '：指定Cellの数式、文字列をクリア

    I1 = 4    'row(4)            ' 貼付先の先頭セルの、行と列 (i1, j1)
    J1 = 2    'column(3) or Columns("C")
    Imax = Imax
    Jmax = Jmax
        With Sheets(DataSheetName)
            .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
                = WorksheetFunction.Transpose(w_Rcurv_base)
        End With

        With Sheets(DataSheetName)
            .Cells(1, J1).Value = "** R_Curvature min /  Parameter = k"
            .Cells(1, J1 + 4).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")
            .Cells(2, J1).Value = "R_Curv min"

            .Cells(I1 - 1, J1).Value = "Phi[rad]"

            For J = 1 To Jmax
                .Cells(2, J1 + J).Value = w_Rcurv_base_min(J)
                .Cells(I1 - 1, J1 + J).Value = " k= " & (0.71 + 0.1 * (J - 1))
            Next J

        End With

End Sub



'========================================================
'-- Data 列貼付
'     １本の曲線の座標点(x､Y)：２列を貼付
'========================================================

Public Sub Paste_curve_data_Num(tmp_Num As Long, Curve_name As String, _
                            curve_x() As Double, curve_y() As Double, DataSheetName As String)

'Public Sub Paste_curve_data(tmp_Num As Long, curve_x() As Double, curve_y() As Double, _
'                              i1 As Long, j1 As Long, DataSheetName As String)
'                (tmp_Num, x_out, y_out, Range("B1").row, Range("B1").Column, DataSheetName)

'----------
    Dim I1 As Long, J1 As Long
    Dim I As Long
    Dim Imax As Long, Jmax As Long
'    Dim tmp_Num As Long, Curve_name As String,
'    Dim curve_x() As Double, curve_y() As Double, DataSheetName As String

'   DataSheetName = "DataSheet_2"
   DataSheetName = "DataSheet_6"    ' Range(cells(i1,j1),cells(i1+dw_n,j1+1))

    Sheets(DataSheetName).Select
    Sheets(DataSheetName).Activate

'-- Data Num
    If tmp_Num < 0 Then      ' Or 9 < tmp_Num Then
     Stop
    End If

'    For I = 0 To 9
            I1 = Range("FE4").Row
            J1 = Range("FE4").Column + (tmp_Num) * 2
'    Next I


'-- Data 貼付

  ' 指定Cellの数式、文字列をクリア
    Sheets(DataSheetName).Range(Cells(I1, J1), Cells(1999, J1 + 1)).ClearContents

  ' 貼付先の先頭セルの、行と列 (i1, j1)
    I1 = I1                         ' i1 = 4    ' Range("A4").row    "4"
    J1 = J1                         ' j1 = 2    ' Range("B1").Column "B"

      With Sheets(DataSheetName)
'          Imax = div_n                          ' Data array length of raws
          Imax = UBound(curve_x)                          ' Data array length of raws
          Jmax = 0                              ' Data array width of collums

        '-- Curve Name
            .Cells(1, J1).Value = "I=[" & Index_I & Format(Now(), "] MM/DD") & Format(Now(), "(HH:mm:ss)")
        '-- Curve Name
            .Cells(2, J1).Value = Curve_name    ' ="Phi_fi"
        '-- Curve X
            .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
                = WorksheetFunction.Transpose(curve_x)
        '-- Curve Y
            .Range(Cells(I1, J1 + 1), Cells(Imax + I1, Jmax + J1 + 1)).Value _
                = WorksheetFunction.Transpose(curve_y)
      End With


End Sub



'========================================================
'-- Data 列貼付
'   Paste_curve_data_Num_2
'     １本の曲線の座標点(x､Y)：２列を貼付
'========================================================

Public Sub Paste_curve_data_Num_2(tmp_Num As Long, Curve_name As String, _
                   curve_x() As Double, curve_y() As Double, DataSheetName As String)

''----------
'    Dim I1 As Long, J1 As Long
'    Dim I As Long
'    Dim Imax As Long, Jmax As Long
''    Dim tmp_Num As Long, Curve_name As String,
''    Dim curve_x() As Double, curve_y() As Double, DataSheetName As String
'
''   DataSheetName = "DataSheet_2"
'   DataSheetName = "DataSheet_7"    ' Range(cells(i1,j1),cells(i1+dw_n,j1+1))
'
'    Sheets(DataSheetName).Select
'    Sheets(DataSheetName).Activate
'
''-- Data Num
'    If tmp_Num < 0 Then
'     Stop
'    End If
'
'          I1 = Range("EV4").Row
'          J1 = Range("EV4").Column + (tmp_Num) * 2
'
''-- Data 貼付 <Paste_curve_data_Num_2>
'
'  ' 指定Cellの数式、文字列をクリア
'    Sheets(DataSheetName).Range(Cells(I1, J1), Cells(1999, J1 + 1)).ClearContents
'
'  ' 貼付先の先頭セルの、行と列 (i1, j1)
'          I1 = I1                         'ex.) i1 = 4    ' Range("A4").row    "4"
'          J1 = J1                         'ex.) j1 = 2    ' Range("B1").Column "B"
'
'      With Sheets(DataSheetName)
'          Imax = div_n                          ' Data array length of raws
'          Jmax = 0                              ' Data array width of collums
'        '-- Curve Name
'            .Cells(2, J1).Value = Curve_name    ' ="Phi_fi"
'        '-- Curve X
'            .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
'                = WorksheetFunction.Transpose(curve_x)
'        '-- Curve Y
'            .Range(Cells(I1, J1 + 1), Cells(Imax + I1, Jmax + J1 + 1)).Value _
'                = WorksheetFunction.Transpose(curve_y)
'      End With

End Sub


'========================================================
'-- Data 列貼付
'   Paste_curve_data_Phi_c_fi
'
'========================================================

Public Sub Paste_curve_data_Phi_c_fi(tmp_cell As String, Curve_name As String, _
                   result_0() As Double, DataSheetName As String)

    Dim the_1 As Double:
'    Dim Curve_name As String

    Dim I As Long, J As Long
    Dim I1 As Long, J1 As Long
    Dim Imax As Long, Jmax As Long

'---------------------------------------
'-- Data 貼付 初期設定

      '--貼付先の先頭セルの、行と列 (i1, j1)

          ' ex.) i1 = 4    ' Range("A4").row    "4"　行
          ' ex.) j1 = 2    ' Range("B1").Column "B"　列
          '       I1 = Range("I4").Row
          '       J1 = Range("I4").Column
          '       Imax = dw_n  '= UBound(Phi_c_fi, 2)   ' Data array length of raws
          '       Jmax = 9     '= UBound(Phi_c_fi, 1)   ' Data array width of collums
          '       ReDim Phi_c_fi(9, dw_n):  ReDim result_0(Jmax, Imax)

        I1 = Range(tmp_cell).Row
        J1 = Range(tmp_cell).Column
        Imax = UBound(Phi_c_fi, 2)    '= dw_n   Data array length of raws
        Jmax = UBound(Phi_c_fi, 1)    '= 9      Data array width of collums

        DataSheetName = "DataSheet_7"

        Sheets(DataSheetName).Select
        Sheets(DataSheetName).Activate

      '--指定Cellの数式、文字列をクリア
        Sheets(DataSheetName).Range(Cells(I1, J1), Cells(1999, J1 + Jmax)).ClearContents


'---------------------------------------
'-- Data 貼付 <Paste_curve_data_Phi_c_fi>

     With Sheets(DataSheetName)

      '-- time stap
        .Cells(1, J1 + 0).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

      '-- Curve Name
        .Cells(3, J1 + 0).Value = "Phi_c_fi(0, I)"   '
        .Cells(3, J1 + 1).Value = "Phi_c_fi(1, I)"   '
        .Cells(3, J1 + 2).Value = "Phi_c_fi(2, I)"   '
        .Cells(3, J1 + 3).Value = "Phi_c_fi(3, I)"   '
        .Cells(3, J1 + 4).Value = "Phi_c_fi(4, I)"   '
        .Cells(3, J1 + 5).Value = "Phi_c_fi(5, I)"   '
        .Cells(3, J1 + 6).Value = "Phi_c_fi(6, I)"   '
        .Cells(3, J1 + 7).Value = "Phi_c_fi(7, I)"   '
        .Cells(3, J1 + 8).Value = "Phi_c_fi(8, I)"   '
        .Cells(3, J1 + 9).Value = "Phi_c_fi(9, I)"   '
    End With

    With Sheets(DataSheetName)
      '-- Curve X
        .Range(Cells(I1, J1), Cells(Imax + I1, Jmax + J1)).Value _
                                = WorksheetFunction.Transpose(Phi_c_fi)

    End With

End Sub



'========================================================
'-- Data 列貼付
'   Paste_curve_data_tg_No
'
'========================================================

Public Sub Paste_curve_data_tg_No(tmp_cell As String, Curve_name As String, _
                   tmp_No As Long, DataSheetName As String)

    Dim the_1 As Double:
    Dim tmp_result()

    Dim I As Long, J As Long
    Dim I1 As Long, J1 As Long
    Dim Imax As Long, Jmax As Long

'---------------------------------------
'-- Data 貼付 初期設定

      '--貼付先の先頭セルの、行と列 (i1, j1)

        I1 = Range(tmp_cell).Row
        J1 = Range(tmp_cell).Column
        Imax = UBound(xg_f, 2)        '= dw_n             'Data array length of raws
        Jmax = 3                      '= UBound(xg_f, 1)  'Data array width of collums

        ReDim tmp_result(2, Imax)

        DataSheetName = "DataSheet_7"

        Sheets(DataSheetName).Select
        Sheets(DataSheetName).Activate

      '--指定Cellの数式、文字列をクリア
        Sheets(DataSheetName).Range(Cells(I1 - 1, J1), Cells(1999, J1 + Jmax - 1)).ClearContents

'---------------------------------------
'-- Data 貼付 <Paste_curve_data_Num_3>

     With Sheets(DataSheetName)

      '-- time stap
        .Cells(1, J1 + 0).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

      '-- Curve Name
        .Cells(3, J1 + 0).Value = Curve_name & "xg_f(" & tmp_No & ", I)"  '
        .Cells(3, J1 + 1).Value = Curve_name & "yg_f(" & tmp_No & ", I)"
        .Cells(3, J1 + 2).Value = Curve_name & "Sg_f(" & tmp_No & ", I)"  '

    End With

    For I = 0 To div_n
         tmp_result(0, I) = xg_f(tmp_No, I)
         tmp_result(1, I) = yg_f(tmp_No, I)
         tmp_result(2, I) = Sg_f(tmp_No, I)
    Next I


    With Sheets(DataSheetName)
      '-- Curve X
        .Range(Cells(I1, J1), Cells(I1 + Imax, J1 + Jmax - 1)).Value _
                                = WorksheetFunction.Transpose(tmp_result)

    End With

End Sub




'========================================================
'-- Data 列貼付
'   Paste_curve_data_tg_No
'
'========================================================

Public Sub Paste_curve_data_tp_No(tmp_cell As String, Curve_name As String, _
                   tmp_No As Long, DataSheetName As String)

    Dim the_1 As Double:
    Dim tmp_result()

    Dim I As Long, J As Long
    Dim I1 As Long, J1 As Long
    Dim Imax As Long, Jmax As Long

'---------------------------------------
'-- Data 貼付 初期設定

      '--貼付先の先頭セルの、行と列 (i1, j1)

        I1 = Range(tmp_cell).Row
        J1 = Range(tmp_cell).Column
        Imax = UBound(xg_f, 2)        '= dw_n             'Data array length of raws
        Jmax = 3                      '= UBound(xg_f, 1)  'Data array width of collums

        ReDim tmp_result(2, Imax)

        DataSheetName = "DataSheet_7"

        Sheets(DataSheetName).Select
        Sheets(DataSheetName).Activate

      '--指定Cellの数式、文字列をクリア
        Sheets(DataSheetName).Range(Cells(I1 - 1, J1), Cells(1999, J1 + Jmax - 1)).ClearContents

'---------------------------------------
'-- Data 貼付 <Paste_curve_data_Num_3>

     With Sheets(DataSheetName)

      '-- time stap
        .Cells(1, J1 + 0).Value = "Date : " & Format(Now(), "yyyy/MM/DD. ") & Format(Now(), "HH:mm:ss")

      '-- Curve Name
        .Cells(3, J1 + 0).Value = Curve_name & "xg_m(" & tmp_No & ", I)"  '
        .Cells(3, J1 + 1).Value = Curve_name & "yg_m(" & tmp_No & ", I)"
        .Cells(3, J1 + 2).Value = Curve_name & "Sg_m(" & tmp_No & ", I)"  '

    End With

    For I = 0 To div_n
         tmp_result(0, I) = xg_m(tmp_No, I)
         tmp_result(1, I) = yg_m(tmp_No, I)
         tmp_result(2, I) = Sg_m(tmp_No, I)
    Next I


    With Sheets(DataSheetName)
      '-- Curve X
        .Range(Cells(I1, J1), Cells(I1 + Imax, J1 + Jmax - 1)).Value _
                                = WorksheetFunction.Transpose(tmp_result)

    End With

End Sub
