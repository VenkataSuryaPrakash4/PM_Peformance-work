*&---------------------------------------------------------------------*
*& Include          ZBDC_PM_PBPLAN_VALIDATION
*&---------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*& Form Upload_logic
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM Upload_logic.

*************************************************
***Converting Excel data into Internal table*****
*************************************************
  CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
    EXPORTING
      i_field_seperator    = 'X'
      i_line_header        = 'X'
      i_tab_raw_data       = gwa_raw_data
      i_filename           = gv_fname
    TABLES
      i_tab_converted_data = gt_excel
    EXCEPTIONS
      conversion_failed    = 1
      OTHERS               = 2.

***Successfull Case, Converting Excel to Internal SAP.
  IF sy-subrc EQ 0.

    DATA: lt_bdc     TYPE TABLE OF bdcdata,
          lt_msg_log TYPE TABLE OF bdcmsgcoll.

***Looping Internal table for processing record by record.
    LOOP AT gt_excel INTO DATA(lwa_excel_tmp).
     DATA(lwa_excel) = lwa_excel_tmp.

***Triggers for every new Equipment number.
      AT NEW equnr.
        lt_bdc = VALUE #( ( program = 'SAPLIWP3' dynpro = '0100' dynbegin = 'X' )
                          ( fnam = 'BDC_CURSOR' fval = 'RMIPM-MPTYP' )
                          ( fnam = 'BDC_OKCODE' fval = '/00' )
                          ( fnam = 'RMIPM-MPTYP' fval = lwa_excel-mptyp )

                          ( program = 'SAPLIWP3' dynpro = '0201' dynbegin = 'X' )
                          ( fnam = 'BDC_OKCODE' fval = 'INIT' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6000SUBSCREEN_HEAD'  )
                          ( fnam = 'RMIPM-WPTXT' fval = lwa_excel-wptxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8002SUBSCREEN_MITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8022SUBSCREEN_BODY2' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6005SUBSCREEN_MAINT_ITEM_TEXT' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWO1                                0100SUBSCREEN_ITEM_1' )
                          ( fnam = 'BDC_CURSOR' fval = 'RIWO1-EQUNR')
                          ( fnam = 'RIWO1-EQUNR' fval = lwa_excel-equnr )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0500SUBSCREEN_ITEM_2')
                          ( fnam = 'RMIPM-AUART' fval = lwa_excel-auart )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6003SUBSCREEN_MLAN_ITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8001SUBSCREEN_MPLAN' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8011SUBSCREEN_BODY1')
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0205SUBSCREEN_CYCLE' )
                          ( fnam = 'RMIPM-ZYKL1' fval = lwa_excel-zykl1 )
                          ( fnam = 'RMIPM-ZEIEH' fval = lwa_excel-zeieh )
                          ( fnam = 'RMIPM-PAK_TEXT' fval = lwa_excel-pak_text )
                          ( fnam = 'RMIPM-POINT' fval = '' )

                          ( program = 'SAPLIWP3' dynpro = '0201' dynbegin = 'X' )
                          ( fnam = 'BDC_OKCODE' fval = '=T\02' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6000SUBSCREEN_HEAD'  )
                          ( fnam = 'BDC_CURSOR' fval = 'RMIPM-WPTXT')
                          ( fnam = 'RMIPM-WPTXT' fval = lwa_excel-wptxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8002SUBSCREEN_MITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8022SUBSCREEN_BODY2' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6005SUBSCREEN_MAINT_ITEM_TEXT' )
                          ( fnam = 'RMIPM-PSTXT' fval = lwa_excel-pstxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWO1                                0100SUBSCREEN_ITEM_1' )
                          ( fnam = 'RIWO1-TPLNR' fval = lwa_excel-tplnr )
                          ( fnam = 'RIWO1-EQUNR' fval = lwa_excel-equnr )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0500SUBSCREEN_ITEM_2')
                          ( fnam = 'RMIPM-IWERK' fval = lwa_excel-iwerk )
                          ( fnam = 'RMIPM-WPGRP' fval = lwa_excel-wpgrp )
                          ( fnam = 'RMIPM-AUART' fval = lwa_excel-auart )
                          ( fnam = 'RMIPM-ILART' fval = lwa_excel-ilart )
                          ( fnam = 'RMIPM-GEWERK' fval = lwa_excel-gewerk )
                          ( fnam = 'RMIPM-WERGW' fval = lwa_excel-wergw )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6003SUBSCREEN_MLAN_ITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8001SUBSCREEN_MPLAN' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8011SUBSCREEN_BODY1')
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0205SUBSCREEN_CYCLE' )
                          ( fnam = 'RMIPM-ZYKL1' fval = lwa_excel-zykl1 )
                          ( fnam = 'RMIPM-ZEIEH' fval = lwa_excel-zeieh )
                          ( fnam = 'RMIPM-PAK_TEXT' fval = lwa_excel-pak_text )
                          ( fnam = 'RMIPM-OFFS1' fval = lwa_excel-offs1 )


                          ( program = 'SAPLIWP3' dynpro = '0201' dynbegin = 'X' )
                          ( fnam = 'BDC_OKCODE' fval = '/00' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6000SUBSCREEN_HEAD'  )
                          ( fnam = 'RMIPM-WPTXT' fval = lwa_excel-wptxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8002SUBSCREEN_MITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8022SUBSCREEN_BODY2' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6005SUBSCREEN_MAINT_ITEM_TEXT' )
                          ( fnam = 'RMIPM-PSTXT' fval = lwa_excel-pstxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWO1                                0100SUBSCREEN_ITEM_1' )
                          ( fnam = 'RIWO1-TPLNR' fval = lwa_excel-tplnr )
                          ( fnam = 'RIWO1-EQUNR' fval = lwa_excel-equnr )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0500SUBSCREEN_ITEM_2')
                          ( fnam = 'RMIPM-IWERK' fval = lwa_excel-iwerk )
                          ( fnam = 'RMIPM-WPGRP' fval = lwa_excel-wpgrp )
                          ( fnam = 'RMIPM-AUART' fval = lwa_excel-auart )
                          ( fnam = 'RMIPM-ILART' fval = lwa_excel-ilart )
                          ( fnam = 'RMIPM-GEWERK' fval = lwa_excel-gewerk )
                          ( fnam = 'RMIPM-WERGW' fval = lwa_excel-wergw )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6003SUBSCREEN_MLAN_ITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8001SUBSCREEN_MPLAN' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8011SUBSCREEN_BODY1')
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0115SUBSCREEN_PARAMETER' )
                          ( fnam = 'BDC_CURSOR' fval = 'RMIPM-ABRHO')
                          ( fnam = 'RMIPM-VSPOS' fval = lwa_excel-vspos )
                          ( fnam = 'RMIPM-HORIZ' fval = lwa_excel-horiz )
                          ( fnam = 'RMIPM-ZEIT' fval = 'X' )
                          ( fnam = 'RMIPM-TOPOS' fval = '100')
                          ( fnam = 'RMIPM-ABRHO' fval = lwa_excel-abrho )
                          ( fnam = 'RMIPM-HUNIT' fval = lwa_excel-hunit )
                          ( fnam = 'RMIPM-VSNEG' fval = lwa_excel-vsneg )
                          ( fnam = 'RMIPM-TONEG' fval = '100')
                          ( fnam = 'RMIPM-SFAKT' fval = lwa_excel-sfakt )

*                          ( program = 'RIPLKO10' dynpro = '1000' dynbegin = 'X')
*                          ( fnam = 'BDC_CURSOR' fval = 'PN_IFLO')
*                          ( fnam = 'BDC_OKCODE' fval = '=ONLI')
*                          ( fnam = 'PN_IFLO' fval = '' )
*                          ( fnam = 'PN_EQUI' fval = '' )
*                          ( fnam = 'PN_IHAN' fval = 'X' )
*                          ( fnam = 'PN_STRNO-LOW' fval = lwa_excel-tplnr )
*                          ( fnam = 'PN_EQUNR-LOW' fval = lwa_excel-equnr )

                          ( program = 'SAPLIWP3' dynpro = '0201' dynbegin = 'X' )
                          ( fnam = 'BDC_OKCODE' fval = '=T\01' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6000SUBSCREEN_HEAD'  )
                          ( fnam = 'RMIPM-WPTXT' fval = lwa_excel-wptxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8002SUBSCREEN_MITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8022SUBSCREEN_BODY2' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6005SUBSCREEN_MAINT_ITEM_TEXT' )
                          ( fnam = 'RMIPM-PSTXT' fval = lwa_excel-pstxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWO1                                0100SUBSCREEN_ITEM_1' )
                          ( fnam = 'RIWO1-TPLNR' fval = lwa_excel-tplnr )
                          ( fnam = 'RIWO1-EQUNR' fval = lwa_excel-equnr )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0500SUBSCREEN_ITEM_2')
                          ( fnam = 'RMIPM-IWERK' fval = lwa_excel-iwerk )
                          ( fnam = 'RMIPM-WPGRP' fval = lwa_excel-wpgrp )
                          ( fnam = 'RMIPM-AUART' fval = lwa_excel-auart )
                          ( fnam = 'RMIPM-ILART' fval = lwa_excel-ilart )
                          ( fnam = 'RMIPM-GEWERK' fval = lwa_excel-gewerk )
                          ( fnam = 'RMIPM-WERGW' fval = lwa_excel-wergw )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6003SUBSCREEN_MLAN_ITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8001SUBSCREEN_MPLAN' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8012SUBSCREEN_BODY1')
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0115SUBSCREEN_PARAMETER' )
                          ( fnam = 'BDC_CURSOR' fval = 'RMIPM-ABRHO')
                          ( fnam = 'RMIPM-VSPOS' fval = lwa_excel-vspos )
                          ( fnam = 'RMIPM-HORIZ' fval = lwa_excel-horiz )
                          ( fnam = 'RMIPM-ZEIT' fval = 'X' )
                          ( fnam = 'RMIPM-TOPOS' fval = '100')
                          ( fnam = 'RMIPM-ABRHO' fval = lwa_excel-abrho )
                          ( fnam = 'RMIPM-HUNIT' fval = lwa_excel-hunit )
                          ( fnam = 'RMIPM-VSNEG' fval = lwa_excel-vsneg )
                          ( fnam = 'RMIPM-TONEG' fval = '100')
                          ( fnam = 'RMIPM-SFAKT' fval = lwa_excel-sfakt )

***Till Here perfect.

                          ( program = 'SAPLIWP3' dynpro = '0201' dynbegin = 'X' )
                          ( fnam = 'BDC_OKCODE' fval = '/00' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6000SUBSCREEN_HEAD'  )
                          ( fnam = 'RMIPM-WPTXT' fval = lwa_excel-wptxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8002SUBSCREEN_MITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8022SUBSCREEN_BODY2' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6005SUBSCREEN_MAINT_ITEM_TEXT' )
                          ( fnam = 'RMIPM-PSTXT' fval = lwa_excel-pstxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWO1                                0100SUBSCREEN_ITEM_1' )
                          ( fnam = 'RIWO1-TPLNR' fval = lwa_excel-tplnr )
                          ( fnam = 'RIWO1-EQUNR' fval = lwa_excel-equnr )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0500SUBSCREEN_ITEM_2')
                          ( fnam = 'RMIPM-IWERK' fval = lwa_excel-iwerk )
                          ( fnam = 'RMIPM-WPGRP' fval = lwa_excel-wpgrp )
                          ( fnam = 'RMIPM-AUART' fval = lwa_excel-auart )
                          ( fnam = 'RMIPM-ILART' fval = lwa_excel-ilart )
                          ( fnam = 'RMIPM-GEWERK' fval = lwa_excel-gewerk )
                          ( fnam = 'RMIPM-WERGW' fval = lwa_excel-wergw )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6003SUBSCREEN_MLAN_ITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8001SUBSCREEN_MPLAN' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8012SUBSCREEN_BODY1')
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0205SUBSCREEN_CYCLE')
                          ( fnam = 'BDC_CURSOR' fval = 'RMIPM-POINT')
                          ( fnam = 'RMIPM-ZYKL1' fval = lwa_excel-zykl1 )
                          ( fnam = 'RMIPM-ZEIEH' fval = lwa_excel-zeieh )
                          ( fnam = 'RMIPM-PAK_TEXT' fval = lwa_excel-pak_text )
                          ( fnam = 'RMIPM-OFFS1' fval = lwa_excel-offs1 )
                          ( fnam = 'RMIPM-POINT' fval = '4' )
***Till Here perfect.

                          ( program = 'SAPLIWP3' dynpro = '0201' dynbegin = 'X' )
                          ( fnam = 'BDC_OKCODE' fval = '=T\02' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6000SUBSCREEN_HEAD'  )
                          ( fnam = 'RMIPM-WPTXT' fval = lwa_excel-wptxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8002SUBSCREEN_MITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8022SUBSCREEN_BODY2' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6005SUBSCREEN_MAINT_ITEM_TEXT' )
                          ( fnam = 'RMIPM-PSTXT' fval = lwa_excel-pstxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWO1                                0100SUBSCREEN_ITEM_1' )
                          ( fnam = 'RIWO1-TPLNR' fval = lwa_excel-tplnr )
                          ( fnam = 'RIWO1-EQUNR' fval = lwa_excel-equnr )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0500SUBSCREEN_ITEM_2')
                          ( fnam = 'RMIPM-IWERK' fval = lwa_excel-iwerk )
                          ( fnam = 'RMIPM-WPGRP' fval = lwa_excel-wpgrp )
                          ( fnam = 'RMIPM-AUART' fval = lwa_excel-auart )
                          ( fnam = 'RMIPM-ILART' fval = lwa_excel-ilart )
                          ( fnam = 'RMIPM-GEWERK' fval = lwa_excel-gewerk )
                          ( fnam = 'RMIPM-WERGW' fval = lwa_excel-wergw )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6003SUBSCREEN_MLAN_ITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8001SUBSCREEN_MPLAN' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8011SUBSCREEN_BODY1')
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0205SUBSCREEN_CYCLE' )
                          ( fnam = 'BDC_CURSOR' fval = 'RMIPM-POINT')
                          ( fnam = 'RMIPM-ZYKL1' fval = lwa_excel-zykl1 )
                          ( fnam = 'RMIPM-ZEIEH' fval = lwa_excel-zeieh )
                          ( fnam = 'RMIPM-PAK_TEXT' fval = lwa_excel-pak_text )
                          ( fnam = 'RMIPM-OFFS1' fval = lwa_excel-offs1 )
                          ( fnam = 'RMIPM-POINT' fval = '4' )

***Till here perfect.

                          ( program = 'SAPLIWP3' dynpro = '0201' dynbegin = 'X' )
                          ( fnam = 'BDC_OKCODE' fval = '/00' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6000SUBSCREEN_HEAD'  )
                          ( fnam = 'RMIPM-WPTXT' fval = lwa_excel-wptxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8002SUBSCREEN_MITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8022SUBSCREEN_BODY2' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6005SUBSCREEN_MAINT_ITEM_TEXT' )
                          ( fnam = 'RMIPM-PSTXT' fval = lwa_excel-pstxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWO1                                0100SUBSCREEN_ITEM_1' )
                          ( fnam = 'RIWO1-TPLNR' fval = lwa_excel-tplnr )
                          ( fnam = 'RIWO1-EQUNR' fval = lwa_excel-equnr )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0500SUBSCREEN_ITEM_2')
                          ( fnam = 'RMIPM-IWERK' fval = lwa_excel-iwerk )
                          ( fnam = 'RMIPM-WPGRP' fval = lwa_excel-wpgrp )
                          ( fnam = 'RMIPM-AUART' fval = lwa_excel-auart )
                          ( fnam = 'RMIPM-ILART' fval = lwa_excel-ilart )
                          ( fnam = 'RMIPM-GEWERK' fval = lwa_excel-gewerk )
                          ( fnam = 'RMIPM-WERGW' fval = lwa_excel-wergw )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6003SUBSCREEN_MLAN_ITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8001SUBSCREEN_MPLAN' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8012SUBSCREEN_BODY1')
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0116SUBSCREEN_PARAMETER' )
                          ( fnam = 'BDC_CURSOR' fval = 'RMIPM-SZAEH')
                          ( fnam = 'RMIPM-VSPOS' fval = lwa_excel-vspos )
                          ( fnam = 'RMIPM-HORIZ' fval = lwa_excel-horiz )
                          ( fnam = 'RMIPM-TOPOS' fval = '100')
                          ( fnam = 'RMIPM-HUNIT' fval = lwa_excel-hunit )
                          ( fnam = 'RMIPM-VSNEG' fval = lwa_excel-vsneg )
                          ( fnam = 'RMIPM-TONEG' fval = '100')
                          ( fnam = 'RMIPM-SFAKT' fval = lwa_excel-sfakt )
                          ( fnam = 'RMIPM-SZAEH' fval = lwa_excel-szaeh )

***Till here perfect.
                          ( program = 'SAPLIWP3' dynpro = '0201' dynbegin = 'X' )
                          ( fnam = 'BDC_OKCODE' fval = '=BU' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6000SUBSCREEN_HEAD'  )
                          ( fnam = 'RMIPM-WPTXT' fval = lwa_excel-wptxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8002SUBSCREEN_MITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8022SUBSCREEN_BODY2' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6005SUBSCREEN_MAINT_ITEM_TEXT' )
                          ( fnam = 'RMIPM-PSTXT' fval = lwa_excel-pstxt )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWO1                                0100SUBSCREEN_ITEM_1' )
                          ( fnam = 'RIWO1-TPLNR' fval = lwa_excel-tplnr )
                          ( fnam = 'RIWO1-EQUNR' fval = lwa_excel-equnr )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0500SUBSCREEN_ITEM_2')
                          ( fnam = 'RMIPM-IWERK' fval = lwa_excel-iwerk )
                          ( fnam = 'RMIPM-WPGRP' fval = lwa_excel-wpgrp )
                          ( fnam = 'RMIPM-AUART' fval = lwa_excel-auart )
                          ( fnam = 'RMIPM-ILART' fval = lwa_excel-ilart )
                          ( fnam = 'RMIPM-GEWERK' fval = lwa_excel-gewerk )
                          ( fnam = 'RMIPM-WERGW' fval = lwa_excel-wergw )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                6003SUBSCREEN_MLAN_ITEM' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8001SUBSCREEN_MPLAN' )
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                8012SUBSCREEN_BODY1')
                          ( fnam = 'BDC_SUBSCR' fval = 'SAPLIWP3                                0116SUBSCREEN_PARAMETER' )
                          ( fnam = 'BDC_CURSOR' fval = 'RMIPM-SZAEH')
                          ( fnam = 'RMIPM-VSPOS' fval = lwa_excel-vspos )
                          ( fnam = 'RMIPM-HORIZ' fval = lwa_excel-horiz )
                          ( fnam = 'RMIPM-TOPOS' fval = '100')
                          ( fnam = 'RMIPM-HUNIT' fval = lwa_excel-hunit )
                          ( fnam = 'RMIPM-VSNEG' fval = lwa_excel-vsneg )
                          ( fnam = 'RMIPM-TONEG' fval = '100')
                          ( fnam = 'RMIPM-SFAKT' fval = lwa_excel-sfakt )
                          ( fnam = 'RMIPM-SZAEH' fval = lwa_excel-szaeh ) ).

      ENDAT.

***Triggers for every last equipment number.
      AT END OF equnr.

        CALL TRANSACTION 'IP41'
                         WITHOUT AUTHORITY-CHECK
                         USING lt_bdc                "Final BDC Internal Table.
                         MODE 'A'                    "No Screens.
                         UPDATE 'S'                  "Synchronous Update.
                         MESSAGES INTO lt_msg_log.   "Message Log.

        CLEAR:lt_bdc.
      ENDAT.

      CLEAR: lwa_excel,lwa_excel_tmp.
    ENDLOOP.

***Failed to convert Excel file to Interal SAP.
  ELSE.
    MESSAGE 'Failed to convert Excel to SAP' TYPE 'E'.
  ENDIF.


ENDFORM.
