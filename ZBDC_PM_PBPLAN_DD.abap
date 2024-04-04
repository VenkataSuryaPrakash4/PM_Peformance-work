*&---------------------------------------------------------------------*
*& Include          ZBDC_PM_PBPLAN_DD
*&---------------------------------------------------------------------*

TYPES: BEGIN OF ty_excel,
         mptyp    TYPE mptyp,
         wptxt    TYPE wptxt,
         pstxt    TYPE postxt,
         tplnr    TYPE tplnr,
         equnr    TYPE equnr,
         iwerk    TYPE iwerk,
         wpgrp    TYPE ingrp,
         auart    TYPE aufart,
         ilart    TYPE ila,
         gewerk   TYPE gewrk,
         wergw    TYPE wergw,
         zykl1    TYPE wzykl1,
         zeieh    TYPE dzeieh,
         pak_text TYPE txzyklus,
         plnty    TYPE plnty,
         plnnr    TYPE plnnr,
         plnal    TYPE plnal,
         offs1    TYPE woffs1,
         vspos    TYPE verschplus,
         vsneg    TYPE verschneg,
         horiz    TYPE horizont,
         abrho    TYPE abrho,
         hunit    TYPE hunit,
         sfakt    TYPE sfakt,
         szaeh    TYPE szaehc,
       END OF ty_excel.

DATA: gt_excel     TYPE TABLE OF ty_excel,
      gv_filename  TYPE ibipparms-path,
      gwa_raw_data TYPE truxs_t_text_data,
      gv_fname     TYPE rlgrap-filename.
