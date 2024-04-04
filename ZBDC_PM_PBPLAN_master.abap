*&---------------------------------------------------------------------*
*& Report ZBDC_PM_PBPLAN
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zbdc_pm_pbplan.

INCLUDE ZBDC_PM_PBPLAN_dd.          "Top Includes.
INCLUDE ZBDC_PM_PBPLAN_ss.          "Selection Screens.
INCLUDE ZBDC_PM_PBPLAN_Validation.  "Upload Logic.


*************************************************
***Enabling F4 functionality for parameter field.
*************************************************

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.

*************************************************
***Enabling Pop-up window, upo clicking on F4****
*************************************************

  CALL FUNCTION 'F4_FILENAME'
    EXPORTING
      program_name  = syst-cprog
      dynpro_number = syst-dynnr
      field_name    = 'P_FILE'
    IMPORTING
      file_name     = gv_filename.

  IF sy-subrc = 0.
    p_file = gv_filename.
    gv_fname = gv_filename.
  ENDIF.

START-OF-SELECTION.
  perform Upload_logic.             "Upload lOgic for BDC.
