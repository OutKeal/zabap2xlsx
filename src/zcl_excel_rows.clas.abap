*----------------------------------------------------------------------*
*       CLASS ZCL_EXCEL_ROWS DEFINITION
*----------------------------------------------------------------------*
*
*----------------------------------------------------------------------*
class ZCL_EXCEL_ROWS definition
  public
  final
  create public .

*"* public components of class ZCL_EXCEL_ROWS
*"* do not include other source files here!!!
*"* protected components of class ZABAP_EXCEL_WORKSHEETS
*"* do not include other source files here!!!
public section.

  methods ADD
    importing
      !IO_ROW type ref to ZCL_EXCEL_ROW .
  methods COPY_ROWS
    importing
      !IP_INDEX type INT4
      !IP_LINES type INT4
      !LO_WORKSHEET type ref to ZCL_EXCEL_WORKSHEET .
  methods CLEAR .
  methods CONSTRUCTOR .
  methods GET
    importing
      !IP_INDEX type I
    returning
      value(EO_ROW) type ref to ZCL_EXCEL_ROW .
  methods GET_ITERATOR
    returning
      value(EO_ITERATOR) type ref to ZCL_EXCEL_COLLECTION_ITERATOR .
  methods IS_EMPTY
    returning
      value(IS_EMPTY) type FLAG .
  methods REMOVE
    importing
      !IO_ROW type ref to ZCL_EXCEL_ROW .
  methods SIZE
    returning
      value(EP_SIZE) type I .
  methods GET_MIN_INDEX
    returning
      value(EP_INDEX) type I .
  methods GET_MAX_INDEX
    returning
      value(EP_INDEX) type I .
  PROTECTED SECTION.
*"* private components of class ZABAP_EXCEL_RANGES
*"* do not include other source files here!!!
private section.

  types:
    BEGIN OF mty_s_hashed_row,
        row_index TYPE int4,
        row       TYPE REF TO zcl_excel_row,
      END OF mty_s_hashed_row .
  types:
    mty_ts_hashed_row TYPE HASHED TABLE OF mty_s_hashed_row WITH UNIQUE KEY row_index .

  data ROWS type ref to ZCL_EXCEL_COLLECTION .
  data ROWS_HASHED type MTY_TS_HASHED_ROW .
ENDCLASS.



CLASS ZCL_EXCEL_ROWS IMPLEMENTATION.


  METHOD add.
    DATA: ls_hashed_row TYPE mty_s_hashed_row.

    ls_hashed_row-row_index = io_row->get_row_index( ).
    ls_hashed_row-row = io_row.

    INSERT ls_hashed_row INTO TABLE rows_hashed.

    rows->add( io_row ).
  ENDMETHOD.                    "ADD


  METHOD clear.
    CLEAR rows_hashed.
    rows->clear( ).
  ENDMETHOD.                    "CLEAR


  METHOD constructor.

    CREATE OBJECT rows.

  ENDMETHOD.                    "CONSTRUCTOR


  METHOD get.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    READ TABLE rows_hashed WITH KEY row_index = ip_index ASSIGNING <ls_hashed_row>.
    IF sy-subrc = 0.
      eo_row = <ls_hashed_row>-row.
    ENDIF.
  ENDMETHOD.                    "GET


  METHOD get_iterator.
    eo_iterator ?= rows->get_iterator( ).
  ENDMETHOD.                    "GET_ITERATOR


  METHOD get_max_index.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    LOOP AT rows_hashed ASSIGNING <ls_hashed_row>.
      IF <ls_hashed_row>-row_index > ep_index.
        ep_index = <ls_hashed_row>-row_index.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD get_min_index.
    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.

    LOOP AT rows_hashed ASSIGNING <ls_hashed_row>.
      IF ep_index = 0 OR <ls_hashed_row>-row_index < ep_index.
        ep_index = <ls_hashed_row>-row_index.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.


  METHOD is_empty.
    is_empty = rows->is_empty( ).
  ENDMETHOD.                    "IS_EMPTY


  METHOD remove.
    DELETE TABLE rows_hashed WITH TABLE KEY row_index = io_row->get_row_index( ) .
    rows->remove( io_row ).
  ENDMETHOD.                    "REMOVE


  METHOD size.
    ep_size = rows->size( ).
  ENDMETHOD.                    "SIZE


  METHOD copy_rows.


    FIELD-SYMBOLS: <ls_hashed_row> TYPE mty_s_hashed_row.
    DATA: ls_hashed_row   TYPE mty_s_hashed_row,
          lo_row          TYPE REF TO zcl_excel_row,
          lo_row_iterator TYPE REF TO zcl_excel_collection_iterator,
          lv_height       TYPE f,
          lv_index        TYPE i.

    CHECK ip_index > 0 AND ip_lines > 0.
    READ TABLE rows_hashed WITH KEY row_index = ip_index ASSIGNING <ls_hashed_row>.
    IF sy-subrc = 0.
      DELETE rows_hashed WHERE row_index > ip_index.
      lv_height = <ls_hashed_row>-row->get_row_height( ).
      lo_row_iterator = me->get_iterator( ).
      WHILE lo_row_iterator->has_next( ) = abap_true.
        lo_row ?= lo_row_iterator->get_next( ).
        lv_index = lo_row->get_row_index( ).

        IF lv_index = ip_index.
          DO ip_lines TIMES.
            CREATE OBJECT lo_row EXPORTING ip_index = ( ip_index + sy-index ).
            IF lv_height > 0.
              lo_row->set_row_height( lv_height ).
            ENDIF.
            me->add( lo_row ).
          ENDDO.
        ENDIF.
        IF lv_index > ip_index.
          lo_row->set_row_index( lv_index + ip_lines ).
          ls_hashed_row-row_index = lv_index + ip_lines.
          ls_hashed_row-row = lo_row.
          INSERT ls_hashed_row INTO TABLE rows_hashed.
        ENDIF.
      ENDWHILE.
    ENDIF.

  ENDMETHOD.
ENDCLASS.
