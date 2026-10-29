*&---------------------------------------------------------------------*
*& Report ZUTI_GPS_IS
*&---------------------------------------------------------------------*
REPORT zuti_gps_is.

TABLES: zgps_appo, zrqti_m2head, sscrfields.


SELECTION-SCREEN BEGIN OF BLOCK main WITH FRAME TITLE TEXT-001.
SELECTION-SCREEN SKIP.
PARAMETERS: rb_fase1 RADIOBUTTON GROUP sel DEFAULT 'X' USER-COMMAND ucomm_main,
            rb_fase2 RADIOBUTTON GROUP sel.
SELECTION-SCREEN SKIP.
SELECTION-SCREEN END OF BLOCK main.

" --- BLOCCO COMPONENTE FASE 1
SELECTION-SCREEN BEGIN OF BLOCK inp_block WITH FRAME TITLE TEXT-104 NO INTERVALS.

" Riga 1: I due Radiobutton vicini
SELECTION-SCREEN BEGIN OF LINE.
SELECTION-SCREEN POSITION 1.
PARAMETERS p_locli RADIOBUTTON GROUP gr1 DEFAULT 'X' USER-COMMAND ucomm_src MODIF ID bl1.
SELECTION-SCREEN COMMENT 3(8) TEXT-004 FOR FIELD p_locli MODIF ID bl1. " Locale

SELECTION-SCREEN POSITION 15.
PARAMETERS p_servi RADIOBUTTON GROUP gr1 MODIF ID bl1.
SELECTION-SCREEN COMMENT 17(8) TEXT-005 FOR FIELD p_servi MODIF ID bl1. " Server
SELECTION-SCREEN END OF LINE.

" Riga 2 e 3 standard
PARAMETERS: p_filel TYPE localfile LOWER CASE MODIF ID ag1 VISIBLE LENGTH 40,
            p_files TYPE string LOWER CASE DEFAULT '\\' MODIF ID ag2 VISIBLE LENGTH 40.

" Riga 4: Checkbox del flag di cancellazione dati
PARAMETERS: p_del TYPE xfeld DEFAULT 'X' AS CHECKBOX MODIF ID bl1.

SELECTION-SCREEN END OF BLOCK inp_block.

" --- BLOCCO COMPONENTE FASE 2
SELECTION-SCREEN BEGIN OF BLOCK avv_block WITH FRAME TITLE TEXT-103 NO INTERVALS.
PARAMETERS: p_anno TYPE zrqti_m2head-anno MODIF ID bl2.
SELECT-OPTIONS: s_qmnum FOR zrqti_m2head-qmnum MODIF ID bl2.
SELECTION-SCREEN END OF BLOCK avv_block.

AT SELECTION-SCREEN OUTPUT.
  LOOP AT SCREEN.

    IF rb_fase1 = 'X' AND screen-group1 = 'BL2'.
      screen-active    = '0'.
      screen-invisible = '1'.
      MODIFY SCREEN.
      CONTINUE.
    ELSEIF rb_fase2 = 'X' AND screen-group1 = 'BL1'.
      screen-active    = '0'.
      screen-invisible = '1'.
      MODIFY SCREEN.
      CONTINUE.
    ENDIF.

    CASE screen-group1.
      WHEN 'AG1'.
        IF rb_fase1 = 'X' AND p_locli = 'X'.
          %_p_filel_%_app_%-text = 'File da locale'.
          screen-active    = '1'.
          screen-invisible = '0'.
        ELSE.
          screen-active    = '0'.
          screen-invisible = '1'.
        ENDIF.
        MODIFY SCREEN.

      WHEN 'AG2'.
        IF rb_fase1 = 'X' AND p_servi = 'X'.
          %_p_files_%_app_%-text = 'File da server'.
          screen-active    = '1'.
          screen-invisible = '0'.
        ELSE.
          screen-active    = '0'.
          screen-invisible = '1'.
        ENDIF.
        MODIFY SCREEN.
    ENDCASE.

  ENDLOOP.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_filel.
  PERFORM call_file_loc.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_files.
  PERFORM call_file_ser.


START-OF-SELECTION.

  " FASE 1: VALORIZZAZIONE TABELLA DI APPOGGIO ZGPS_APPO
  IF rb_fase1 = 'X'.

    IF ( p_locli = 'X' AND p_filel IS INITIAL ) OR ( p_servi = 'X' AND p_files IS INITIAL ).
      MESSAGE 'Inserire il percorso del file di input.' TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    IF p_del = 'X'. "flag attivo
      DELETE FROM zgps_appo.
      COMMIT WORK.
    ENDIF.

    TYPES: BEGIN OF ty_excel_row,
             globalid    TYPE string, " Colonna A
             ogg_all     TYPE string, " Colonna B (OGG_ALL)
             impianto    TYPE string, " Colonna C
             codicevia   TYPE string, " Colonna D
             comune      TYPE string, " Colonna E
             x_epsg      TYPE string, " Colonna F
             y_epsg      TYPE string, " Colonna G
             tiporete    TYPE string, " Colonna H
             latitudine  TYPE string, " Colonna I (Coordinata Y)
             longitudine TYPE string, " Colonna J (Coordinata X)
           END OF ty_excel_row.

    DATA: lt_excel_data     TYPE STANDARD TABLE OF ty_excel_row,
          ls_excel_row      TYPE ty_excel_row,
          lv_file_converted TYPE rlgrap-filename,
          lt_raw_data       TYPE truxs_t_text_data,
          lv_righe_lette    TYPE i,
          lv_line           TYPE string,
          lv_is_header      TYPE c VALUE 'X',
          lt_csv_tokens     TYPE STANDARD TABLE OF string.

    " --- CARICAMENTO DA LOCALE PC UTENTE
    IF p_locli = 'X'.
      lv_file_converted = p_filel.


      CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
        EXPORTING
          i_line_header        = 'X'
          i_filename           = lv_file_converted
          i_tab_raw_data       = lt_raw_data
        TABLES
          i_tab_converted_data = lt_excel_data
        EXCEPTIONS
          conversion_failed    = 1
          OTHERS               = 2.

      IF sy-subrc <> 0.
        MESSAGE 'Errore durante la conversione o lettura del file Excel locale.' TYPE 'I' DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.


      " ---------------------
      " ---------------------

      " caricamento da server
    ELSEIF p_servi = 'X'.
      DATA: lv_server_path TYPE string,
            lv_content     TYPE xstring,
            lv_buffer      TYPE xstring,
            lt_binary_tab  TYPE STANDARD TABLE OF x255,
            lv_file_length TYPE i,
            lv_temp_local  TYPE string,
            lv_del_rc      TYPE i.

      lv_server_path = p_files.

      OPEN DATASET lv_server_path FOR INPUT IN BINARY MODE.
      IF sy-subrc <> 0.
        MESSAGE |Impossibile aprire il file Excel nel percorso server: { p_files }| TYPE 'I' DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.

      CLEAR lv_content.
      DO.
        READ DATASET lv_server_path INTO lv_buffer.
        IF sy-subrc <> 0.
          CONCATENATE lv_content lv_buffer INTO lv_content IN BYTE MODE.
          EXIT.
        ENDIF.
        CONCATENATE lv_content lv_buffer INTO lv_content IN BYTE MODE.
      ENDDO.
      CLOSE DATASET lv_server_path.

      IF sy-batch = 'X'.
        " REPORT GIRA IN BACKGROUND
        DATA: lo_excel_batch  TYPE REF TO cl_fdt_xl_spreadsheet,
              lt_worksheets_b TYPE string_table,
              lo_data_batch   TYPE REF TO data.

        FIELD-SYMBOLS: <lt_sheet_batch>  TYPE STANDARD TABLE,
                       <ls_sheet_row_b>  TYPE any,
                       <lv_cell_value_b> TYPE any.

        TRY.
            CREATE OBJECT lo_excel_batch
              EXPORTING
                document_name = p_files
                xdocument     = lv_content.
          CATCH cx_fdt_excel_core.
            RETURN.
        ENDTRY.

        lo_excel_batch->if_fdt_doc_spreadsheet~get_worksheet_names( IMPORTING worksheet_names = lt_worksheets_b ).
        READ TABLE lt_worksheets_b INTO DATA(lv_sheet_batch) INDEX 2. " N.B. prende IL 2TAB del file excel

        IF sy-subrc = 0.
          lo_data_batch = lo_excel_batch->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
                            worksheet_name = lv_sheet_batch ).

          ASSIGN lo_data_batch->* TO <lt_sheet_batch>.

          LOOP AT <lt_sheet_batch> ASSIGNING <ls_sheet_row_b>.
            IF sy-tabix = 1. CONTINUE. ENDIF.
            CLEAR ls_excel_row.

            " Colonna B: Oggetto di Allacciamento
            ASSIGN COMPONENT 2 OF STRUCTURE <ls_sheet_row_b> TO <lv_cell_value_b>.
            IF sy-subrc = 0 AND <lv_cell_value_b> IS ASSIGNED. ls_excel_row-ogg_all = <lv_cell_value_b>. ENDIF.

            " Colonna F: Coordinata X
            ASSIGN COMPONENT 6 OF STRUCTURE <ls_sheet_row_b> TO <lv_cell_value_b>.
            IF sy-subrc = 0 AND <lv_cell_value_b> IS ASSIGNED. ls_excel_row-x_epsg = <lv_cell_value_b>. ENDIF.

            " Colonna G: Coordinata Y
            ASSIGN COMPONENT 7 OF STRUCTURE <ls_sheet_row_b> TO <lv_cell_value_b>.
            IF sy-subrc = 0 AND <lv_cell_value_b> IS ASSIGNED. ls_excel_row-y_epsg = <lv_cell_value_b>. ENDIF.

            IF ls_excel_row-ogg_all IS NOT INITIAL.
              APPEND ls_excel_row TO lt_excel_data.
            ENDIF.
          ENDLOOP.
        ENDIF.

      ELSE.
        " REPORT GIRA A SCHERMO (NO BACKGROUND)
        CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
          EXPORTING
            buffer        = lv_content
          IMPORTING
            output_length = lv_file_length
          TABLES
            binary_tab    = lt_binary_tab.

        cl_gui_frontend_services=>get_temp_directory( CHANGING temp_dir = lv_temp_local ).
        cl_gui_cfw=>flush( ).

        lv_temp_local = |{ lv_temp_local }\\temp_cartografia_{ sy-uname }.xlsx|.

        cl_gui_frontend_services=>gui_download(
          EXPORTING
            bin_filesize              = lv_file_length
            filename                  = lv_temp_local
            filetype                  = 'BIN'
          CHANGING
            data_tab                  = lt_binary_tab
          EXCEPTIONS
            file_write_error          = 1
            OTHERS                    = 2 ).

        IF sy-subrc = 0.
          lv_file_converted = lv_temp_local.

          CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
            EXPORTING
              i_line_header        = 'X'
              i_filename           = lv_file_converted
              i_tab_raw_data       = lt_raw_data
            TABLES
              i_tab_converted_data = lt_excel_data
            EXCEPTIONS
              conversion_failed    = 1
              OTHERS               = 2.

          cl_gui_frontend_services=>file_delete(
            EXPORTING
              filename = lv_temp_local
            CHANGING
              rc       = lv_del_rc
            EXCEPTIONS
              OTHERS   = 1 ).

          IF sy-batch = abap_false.
            IF sy-subrc <> 0 OR lt_excel_data IS INITIAL.
              MESSAGE 'Errore durante la conversione del file Excel.' TYPE 'I' DISPLAY LIKE 'E'.
              RETURN.
            ENDIF.
          ELSE.
            MESSAGE 'Impossibile trasferire il file temporaneo dal Server al PC locale.' TYPE 'I' DISPLAY LIKE 'E'.
            RETURN.
          ENDIF.
        ENDIF.
      ENDIF.
    ENDIF.


    "-------------------------
    "-------------------------

    lv_righe_lette = lines( lt_excel_data ).

    DATA: lt_gps_appo TYPE STANDARD TABLE OF zgps_appo,
          ls_gps_appo TYPE zgps_appo,
          ls_address  TYPE ieadrc,
          lv_tabix    TYPE string,
          lv_coord_x  TYPE zgps_appo-zzgeointx,
          lv_coord_y  TYPE zgps_appo-zzgeointy.

    LOOP AT lt_excel_data INTO ls_excel_row.

      IF sy-tabix MOD 50 = 0.
        lv_tabix = sy-tabix.
        CONDENSE lv_tabix.
        CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
          EXPORTING
            percentage = 0
            text       = |Elaborazione indirizzi: riga { lv_tabix } di { lv_righe_lette }|.
      ENDIF.

      IF ls_excel_row-ogg_all IS INITIAL.
        CONTINUE.
      ENDIF.

      CLEAR: ls_gps_appo, ls_address.
      ls_gps_appo-mandt = sy-mandt.
      ls_gps_appo-haus  = ls_excel_row-ogg_all.
      CONDENSE ls_gps_appo-haus.

      CALL FUNCTION 'ISU_ADDRESS_PROVIDE'
        EXPORTING
          x_address_type  = 'C'
          x_haus          = ls_gps_appo-haus
        IMPORTING
          y_ieadrc        = ls_address
        EXCEPTIONS
          not_found       = 1
          parameter_error = 2
          system_error    = 3
          OTHERS          = 4.

      IF sy-subrc = 0.
        READ TABLE ls_address INTO DATA(ls_address_row) INDEX 1.
        IF sy-subrc = 0.
          ls_gps_appo-city1      = ls_address_row-city1.       " Località
          ls_gps_appo-street     = ls_address_row-street.      " Via
          ls_gps_appo-house_num1 = ls_address_row-house_num1.  " Civico
          ls_gps_appo-house_num2 = ls_address_row-str_suppl1.  " Estensione civico
        ENDIF.
      ENDIF.

*      DATA(lv_tmp_x) = ls_excel_row-longitudine. "colonna j
*      DATA(lv_tmp_y) = ls_excel_row-latitudine. " colonna i

      DATA(lv_tmp_x) = ls_excel_row-x_epsg. " Prende la colonna F invece della J
      DATA(lv_tmp_y) = ls_excel_row-y_epsg. " Prende la colonna G invece della I

      REPLACE ALL OCCURRENCES OF '"' IN lv_tmp_x WITH ''.
      REPLACE ALL OCCURRENCES OF '"' IN lv_tmp_y WITH ''.
      REPLACE ALL OCCURRENCES OF ',' IN lv_tmp_x WITH '.'.
      REPLACE ALL OCCURRENCES OF ',' IN lv_tmp_y WITH '.'.
      CONDENSE lv_tmp_x.
      CONDENSE lv_tmp_y.

      lv_coord_x = lv_tmp_x.
      lv_coord_y = lv_tmp_y.

      ls_gps_appo-zzgeointx = lv_coord_x.
      ls_gps_appo-zzgeointy = lv_coord_y.


      APPEND ls_gps_appo TO lt_gps_appo.
    ENDLOOP.

    IF lt_gps_appo IS NOT INITIAL.
      INSERT zgps_appo FROM TABLE lt_gps_appo ACCEPTING DUPLICATE KEYS.
      COMMIT WORK.
      MESSAGE |Fase 1 completata! Inseriti { sy-dbcnt } record nella tabella ZGPS_APPO.| TYPE 'I'.
    ELSE.
      MESSAGE 'Nessun dato valido elaborato da inserire.' TYPE 'I' DISPLAY LIKE 'W'.
    ENDIF.

    " ESECUZIONE FASE 2: VALORIZZAZIONE AVVISI IS
  ELSEIF rb_fase2 = 'X'.

    " Controllo iniziale bloccante se la tabella di appoggio è vuota
    DATA: lv_count_appo TYPE i.
    SELECT COUNT(*)
      FROM zgps_appo
    INTO lv_count_appo.

    IF lv_count_appo = 0.
      MESSAGE 'Devi prima eseguire il passo 1: tabella ZGPS_APPO vuota' TYPE 'I'.
      RETURN.
    ENDIF.

    " Controllo Anno obbligatorio
    IF p_anno IS INITIAL.
      MESSAGE 'L''anno di selezione è obbligatorio.' TYPE 'S' DISPLAY LIKE 'E'.
      RETURN.
    ENDIF.

    TYPES: BEGIN OF ty_avvisi,
             anno    TYPE zrqti_m2head-anno,
             qmnum   TYPE zrqti_m2head-qmnum,
             iloan   TYPE viqmel-iloan,
             tplnr   TYPE viqmel-tplnr,
             geointx TYPE viqmel-zzgeointx,
           END OF ty_avvisi.

    DATA: lt_avvisi TYPE STANDARD TABLE OF ty_avvisi.

    " Recupero avvisi accedendo a ZRQTI_M2HEAD in INNER JOIN con VIQMEL
    SELECT a~anno, a~qmnum, v~iloan, v~tplnr, v~zzgeointx
      FROM zrqti_m2head AS a
      INNER JOIN viqmel AS v ON v~qmnum = a~qmnum
      INTO TABLE @lt_avvisi
      WHERE a~anno  = @p_anno
    AND a~qmnum IN @s_qmnum.

    IF lt_avvisi IS INITIAL.
      MESSAGE 'Nessun avviso trovato nella tabella ZRQTI_M2HEAD per i criteri inseriti.' TYPE 'I'.
      RETURN.
    ENDIF.

    " Se l'avviso ha già le coordinate GPS valorizzate lo si salta
    DELETE lt_avvisi WHERE geointx IS NOT INITIAL.

    IF lt_avvisi IS INITIAL.
      MESSAGE 'Tutti gli avvisi selezionati hanno già le coordinate GPS valorizzate.' TYPE 'I'.
      RETURN.
    ENDIF.

    TYPES: BEGIN OF ty_log,
             qmnum      TYPE viqmel-qmnum,
             city1      TYPE adrc-city1,
             street     TYPE adrc-street,
             house_num1 TYPE adrc-house_num1,
             house_num2 TYPE adrc-house_num2,
             zzgeointx  TYPE viqmel-zzgeointx,
             zzgeointy  TYPE viqmel-zzgeointy,
             status     TYPE string,
           END OF ty_log.

    DATA: lt_log        TYPE STANDARD TABLE OF ty_log,
          ls_log        TYPE ty_log,
          ls_iloa       TYPE iloa,
          ls_adrc       TYPE adrc,
          ls_appo       TYPE zgps_appo,
          lv_aggiornati TYPE i VALUE 0,
          lv_tot_selez  TYPE i.

    lv_tot_selez = lines( lt_avvisi ).

    " Elaborazione ad 1 ad 1 degli avvisi in selezione
    LOOP AT lt_avvisi INTO DATA(ls_avviso).
      CLEAR: ls_iloa, ls_adrc, ls_appo, ls_log.
      ls_log-qmnum = ls_avviso-qmnum.

      " Accesso alla tabella ILOA tramite ubicazione e sede tecnica per trovare il codice indirizzo
      SELECT SINGLE adrnr
        FROM iloa
        INTO @ls_iloa-adrnr
        WHERE iloan = @ls_avviso-iloan
      AND tplnr = @ls_avviso-tplnr.

      IF sy-subrc <> 0 OR ls_iloa-adrnr IS INITIAL.
        ls_log-status = 'Ubicazione / Indirizzo ILOA non trovato'.
        APPEND ls_log TO lt_log.
        CONTINUE.
      ENDIF.

      " Accesso alla tabella ADRC tramite il codice indirizzo
      SELECT SINGLE city1, street, house_num1, str_suppl1 AS house_num2
        FROM adrc
        INTO CORRESPONDING FIELDS OF @ls_adrc
      WHERE addrnumber = @ls_iloa-adrnr.


      IF sy-subrc <> 0.
        ls_log-status = 'Record indirizzo ADRC non trovato'.
        APPEND ls_log TO lt_log.
        CONTINUE.
      ENDIF.

      ls_log-city1      = ls_adrc-city1.
      ls_log-street     = ls_adrc-street.
      ls_log-house_num1 = ls_adrc-house_num1.
      ls_log-house_num2 = ls_adrc-house_num2.

      " Verifica corrispondenza per esteso (località, via, civico, estensione civico) su ZGPS_APPO
      SELECT SINGLE zzgeointx, zzgeointy
        FROM zgps_appo
        INTO CORRESPONDING FIELDS OF @ls_appo
        WHERE city1      = @ls_adrc-city1
          AND street     = @ls_adrc-street
          AND house_num1 = @ls_adrc-house_num1
      AND house_num2 = @ls_adrc-house_num2.

      IF sy-subrc = 0.
        ls_log-zzgeointx = ls_appo-zzgeointx.
        ls_log-zzgeointy = ls_appo-zzgeointy.
        ls_log-status    = 'Accoppiato - GPS Trovato'.

        UPDATE qmel
           SET zzgeointx = @ls_appo-zzgeointx,
               zzgeointy = @ls_appo-zzgeointy
         WHERE qmnum     = @ls_avviso-qmnum.

        UPDATE zrqti_m2head
           SET origx = @ls_appo-zzgeointx,
               origy = @ls_appo-zzgeointy
         WHERE qmnum = @ls_avviso-qmnum
           AND anno  = @ls_avviso-anno.

        lv_aggiornati = lv_aggiornati + 1.
      ELSE.
        ls_log-status = 'Nessuna corrispondenza in tabella appoggio'.
        CLEAR: ls_log-zzgeointx, ls_log-zzgeointy.
      ENDIF.

      APPEND ls_log TO lt_log.


    ENDLOOP.

    IF lv_aggiornati > 0.
      COMMIT WORK.
    ENDIF.

    DATA: lo_alv  TYPE REF TO cl_salv_table,
          lo_cols TYPE REF TO cl_salv_columns_table,
          lo_col  TYPE REF TO cl_salv_column_table.

    TRY.
        cl_salv_table=>factory(
          IMPORTING
            r_salv_table = lo_alv
          CHANGING
            t_table      = lt_log ).

        lo_alv->get_functions( )->set_all( abap_true ).

        lo_cols = lo_alv->get_columns( ).

        lo_col ?= lo_cols->get_column( 'QMNUM' ).
        lo_col->set_long_text( 'Numero Avviso IS' ).
        lo_col->set_medium_text( 'Num. Avviso' ).

        lo_col ?= lo_cols->get_column( 'CITY1' ).
        lo_col->set_long_text( 'Località / Comune' ).
        lo_col->set_medium_text( 'Località' ).

        lo_col ?= lo_cols->get_column( 'STREET' ).
        lo_col->set_long_text( 'Via / Piazza' ).
        lo_col->set_medium_text( 'Via' ).

        lo_col ?= lo_cols->get_column( 'HOUSE_NUM1' ).
        lo_col->set_long_text( 'Numero Civico' ).
        lo_col->set_medium_text( 'Civico' ).

        lo_col ?= lo_cols->get_column( 'HOUSE_NUM2' ).
        lo_col->set_long_text( 'Estensione Civico' ).
        lo_col->set_medium_text( 'Estensione' ).

        lo_col ?= lo_cols->get_column( 'ZZGEOINTX' ).
        lo_col->set_long_text( 'Coordinata X (Longitudine)' ).
        lo_col->set_medium_text( 'Coord. X' ).

        lo_col ?= lo_cols->get_column( 'ZZGEOINTY' ).
        lo_col->set_long_text( 'Coordinata Y (Latitudine)' ).
        lo_col->set_medium_text( 'Coord. Y' ).

        lo_col ?= lo_cols->get_column( 'STATUS' ).
        lo_col->set_long_text( 'Esito del Match GPS' ).
        lo_col->set_medium_text( 'Esito Match' ).

        " Mostra la griglia ALV a schermo
        lo_alv->display( ).

      CATCH cx_salv_msg.
        MESSAGE 'Errore durante la generazione della griglia ALV.' TYPE 'I' DISPLAY LIKE 'E'.
    ENDTRY.

    "MESSAGE |Elaborazione conclusa. Aggiornati { lv_aggiornati } avvisi su { lv_tot_selez } elaborati.| TYPE 'I'.

  ENDIF.



*
*    " log
*    WRITE: /
*    ULINE.
*    WRITE: /3 'Num. Avviso', 16 'Località', 38 'Via', 65 'Civ.', 70 'Est.', 77 'Coordinata X (Long)', 98 'Coordinata Y (Lat)', 120 'Esito Match'.
*    ULINE.
*
*    LOOP AT lt_log INTO ls_log.
*      WRITE: / ls_log-qmnum UNDER 'Num. Avviso',
*               ls_log-city1 UNDER 'Località',
*               ls_log-street UNDER 'Via',
*               ls_log-house_num1 UNDER 'Civ.',
*               ls_log-house_num2 UNDER 'Est.',
*               ls_log-zzgeointx UNDER 'Coordinata X (Long)',
*               ls_log-zzgeointy UNDER 'Coordinata Y (Lat)',
*               ls_log-status UNDER 'Esito Match'.
*    ENDLOOP.
*    ULINE.

  "MESSAGE |Elaborazione conclusa. Aggiornati { lv_aggiornati } avvisi su { lv_tot_selez } elaborati.| TYPE 'I'.


*&---------------------------------------------------------------------*
*&      Form  CALL_FILE_LOC
*&---------------------------------------------------------------------*
FORM call_file_loc.
  DATA: lv_rc   TYPE sy-subrc,
        lt_file TYPE filetable,
        ls_file LIKE LINE OF lt_file.

  ls_file-filename = p_filel.
  APPEND ls_file TO lt_file.

  cl_gui_frontend_services=>file_open_dialog(
    CHANGING
      file_table = lt_file
      rc         = lv_rc ).

  READ TABLE lt_file INTO ls_file INDEX 1.
  IF sy-subrc = 0.
    p_filel = ls_file-filename.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  CALL_FILE_SER
*&---------------------------------------------------------------------*
FORM call_file_ser.
  DATA: lv_server_file TYPE string.
  lv_server_file = p_files.

  " Chiamata alla funzione aziendale per esplorare il server AL11
  CALL FUNCTION '/SAPDMC/LSM_F4_SERVER_FILE'
    IMPORTING
      serverfile       = lv_server_file
    EXCEPTIONS
      canceled_by_user = 1
      OTHERS           = 2.

  IF sy-subrc = 0 AND lv_server_file IS NOT INITIAL.
    p_files = lv_server_file.
  ENDIF.
ENDFORM.