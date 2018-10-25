package com.penghaohuan.excel.model;

/**
 * Description of the position of the Excel cell.
 *
 * @author penghaohuan
 */
public class CellPosition {

    /**
     * 行.
     */
    private int row;

    /**
     * 列.
     */
    private int column;

    public CellPosition() {
    }

    public CellPosition(int row, int column) {
        this.row = row;
        this.column = column;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getColumn() {
        return column;
    }

    public void setColumn(int column) {
        this.column = column;
    }

}
