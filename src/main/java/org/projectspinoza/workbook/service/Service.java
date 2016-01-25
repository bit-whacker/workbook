package org.projectspinoza.workbook.service;

public interface Service {

    public void read();
    public void write();
    /**
     * updates specific column in the workbook sheet
     * 
     * @param sheet
     * @param row
     * @param column
     * @param value
     * @throws Exception
     */
    public void update(int sheet, int row, int column, double value) throws Exception;
    public void delete();
}
