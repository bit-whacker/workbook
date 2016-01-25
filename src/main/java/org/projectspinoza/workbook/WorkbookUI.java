package org.projectspinoza.workbook;

import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.projectspinoza.workbook.service.WorkbookService;

public class WorkbookUI {
    
    private final int irr2015 = 0;
    private final int irc2015 = 1;
    private final int irr2016 = 0;
    private final int irc2016 = 2;
    private final int irr2017 = 0;
    private final int irc2017 = 3;
    
    
    private JFrame frame;
    private JPanel firstTablePanel;
    private JScrollPane scrollPane;
    private JTable table;
    private JLabel lblFormula;
    private JLabel lblNewLabel;
    private JPanel panel;
    private JScrollPane scrollPane_1;
    private JTable table_1;
    private JButton btnNewButton;
    private WorkbookService workbookService;
    
    public static WorkbookUI window;
    
    public WorkbookUI() {
        initialize();
    }

    /**
     * Launch the application.
     */
    public static void main(String[] args) {
        EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    window = new WorkbookUI();
                    window.frame.setVisible(true);
                    window.syncTableWithWorkbook();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }

    /**
     * Initialize the contents of the frame.
     */
    private void initialize() {
        frame = new JFrame("Consolidated Workbook - v1.0.0");
        frame.setResizable(false);
        frame.setBounds(100, 100, 704, 480);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.getContentPane().setLayout(null);
        frame.getContentPane().add(getFirstTablePanel());
        frame.getContentPane().add(getLblFormula());
        frame.getContentPane().add(getLblNewLabel());
        frame.getContentPane().add(getPanel());
        frame.getContentPane().add(getBtnNewButton());
        workbookService = new WorkbookService();
    }

    private JPanel getFirstTablePanel() {
        if (firstTablePanel == null) {
            firstTablePanel = new JPanel();
            firstTablePanel.setBounds(10, 35, 688, 213);
            firstTablePanel.setLayout(null);
            firstTablePanel.add(getScrollPane());
        }
        return firstTablePanel;
    }

    private JScrollPane getScrollPane() {
        if (scrollPane == null) {
            scrollPane = new JScrollPane();
            scrollPane.setBounds(0, 11, 688, 191);
            scrollPane.setViewportView(getTable());
        }
        return scrollPane;
    }

    private JTable getTable() {
        if (table == null) {
            table = new JTable();
            table.setModel(new DefaultTableModel(
                    new Object[][] { 
                        { "Revenue", null, null, null, null }, { "Cost of Sales", null, null, null, null },
                        { "Gorss Margin", null, null, null, null }, { "SG&A", null, null, null, null },
                        { "Interest", null, null, null, null },
                        { "Income Before Taxes - Scenario", null, null, null, null },
                        { "Income Before Taxes - Base Case", null, null, null, null },
                        { "Net Change - $m", null, null, null, null },
                        { "Net Change In Interest Rate - %", null, null, null, null }, },
                    new String[] { "(All amounts are in $ millions except interest rates)", "2015", "2016", "2017",
                            "Total" }) {
                boolean[] columnEditables = new boolean[] { false, false, false, false, false };

                public boolean isCellEditable(int row, int column) {
                    return columnEditables[column];
                }
            });
            table.getColumnModel().getColumn(0).setResizable(false);
            table.getColumnModel().getColumn(0).setPreferredWidth(311);
            table.getColumnModel().getColumn(1).setPreferredWidth(86);
        }
        return table;
    }

    private JLabel getLblFormula() {
        if (lblFormula == null) {
            lblFormula = new JLabel("Formula");
            lblFormula.setBounds(10, 10, 72, 25);
        }
        return lblFormula;
    }

    private JLabel getLblNewLabel() {
        if (lblNewLabel == null) {
            lblNewLabel = new JLabel("--");
            lblNewLabel.setBounds(71, 11, 550, 24);
        }
        return lblNewLabel;
    }

    private JPanel getPanel() {
        if (panel == null) {
            panel = new JPanel();
            panel.setBounds(10, 270, 668, 95);
            panel.setLayout(null);
            panel.add(getScrollPane_1());
        }
        return panel;
    }

    private JScrollPane getScrollPane_1() {
        if (scrollPane_1 == null) {
            scrollPane_1 = new JScrollPane();
            scrollPane_1.setBounds(10, 11, 648, 73);
            scrollPane_1.setViewportView(getTable_1());
        }
        return scrollPane_1;
    }

    private JTable getTable_1() {
        if (table_1 == null) {
            table_1 = new JTable();
            table_1.setModel(new DefaultTableModel(
                    new Object[][] { 
                        { "New Interest Rate - %", null, null, null },
                        { "Base Case Interest Rate - %", null, null, null }, },
                    new String[] { "Interest Rate Assumptions", "2015", "2016", "2017" }) {
                boolean[] columnEditables = new boolean[] { false, true, true, true };

                public boolean isCellEditable(int row, int column) {
                    return columnEditables[column];
                }
            });
            table_1.getColumnModel().getColumn(0).setResizable(false);
            table_1.getColumnModel().getColumn(0).setPreferredWidth(262);
            table_1.getColumnModel().getColumn(1).setPreferredWidth(114);
            table_1.getColumnModel().getColumn(2).setPreferredWidth(103);
            table_1.getColumnModel().getColumn(3).setPreferredWidth(104);
        }
        return table_1;
    }

    private JButton getBtnNewButton() {
        if (btnNewButton == null) {
            btnNewButton = new JButton("Update");
            btnNewButton.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent arg0) {
                    // . TODO on action perform
                    String nirr2015 = getTable_1().getValueAt(irr2015, irc2015).toString();
                    String nirr2016 = getTable_1().getValueAt(irr2016, irc2016).toString();
                    String nirr2017 = getTable_1().getValueAt(irr2017, irc2017).toString();
                    
                    String birr2015 = getTable_1().getValueAt(irr2015+1, irc2015).toString();
                    String birr2016 = getTable_1().getValueAt(irr2016+1, irc2016).toString();
                    String birr2017 = getTable_1().getValueAt(irr2017+1, irc2017).toString();
                    
                    String formatNirr = "New Interest Rate{ 2015["+nirr2015+"], 2016["+nirr2016+"], 2017["+nirr2017+"]}";
                    String formatBirr = "Base Case Interest Rate{ 2015["+birr2015+"], 2016["+birr2016+"], 2017["+birr2017+"]}";
                    
                    
                    System.out.println(formatNirr);
                    System.out.println(formatBirr);
                    try{
                        Double nirr2015d = Double.parseDouble(nirr2015);
                        nirr2015d = nirr2015d/100;
                        nirr2015d = round(nirr2015d, 4);
                   
                        Double nirr2016d = Double.parseDouble(nirr2016);
                        nirr2016d = nirr2016d/100;
                        nirr2016d = round(nirr2016d, 4);
                        
                        Double nirr2017d = Double.parseDouble(nirr2017);
                        nirr2017d = nirr2017d/100;
                        nirr2017d = round(nirr2017d, 4);
                        
                        ////////////////////////////////////////////////////////////////
                        
                        Double birr2015d = Double.parseDouble(birr2015);
                        birr2015d = birr2015d/100;
                        birr2015d = round(birr2015d, 4);
                   
                        Double birr2016d = Double.parseDouble(birr2016);
                        birr2016d = birr2016d/100;
                        birr2016d = round(birr2016d, 4);
                        
                        Double birr2017d = Double.parseDouble(birr2017);
                        birr2017d = birr2017d/100;
                        birr2017d = round(birr2017d, 4);
                        
                        
                        window.getWorkbookService().update(0, irr2015+19, irc2015+1 , nirr2015d);
                        window.getWorkbookService().update(0, irr2016+19, irc2016+1 , nirr2016d);
                        window.getWorkbookService().update(0, irr2017+19, irc2017+1 , nirr2017d);
                        
                        window.getWorkbookService().update(0, irr2015+20, irc2015+1 , birr2015d);
                        window.getWorkbookService().update(0, irr2016+20, irc2016+1 , birr2016d);
                        window.getWorkbookService().update(0, irr2017+20, irc2017+1 , birr2017d);
                        
                        
                        window.syncTableWithWorkbook();
                    }catch(Exception ex){
                        System.err.println("got exception during value updation!!!");
                    }
                }
            });
            btnNewButton.setBounds(476, 370, 191, 23);
        }
        return btnNewButton;
    }
    
    public void syncTableWithWorkbook() throws IOException, FileNotFoundException{
        FileInputStream fis = new FileInputStream(new File("workbook.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook (fis);
        XSSFSheet xlsheet = workbook.getSheetAt(0);
        
//        rows -> 7:20 
//        columns -> 2:5
        
        for(int row = 7; row <= 20; row++){
            if(row == 13 || row == 17 || row == 18){continue;}
            XSSFRow xlrow = xlsheet.getRow(row);
            for(int cell = 2; cell <= 5; cell++){
                XSSFCell xlcell = xlrow.getCell(cell);
                if(xlcell == null){
                    continue;
                }
                Double v = null;
                try{
                    v = Double.parseDouble(xlcell.getRawValue());
                    v = round(v, 4);
                }catch(Exception ex){
                   System.err.println("cannot convert to double: " + v); 
                }
                updateCell(row, cell, v);
                System.out.println(row+"x"+cell+" => " + v);
            }
        }
        
        fis.close();
    }
    public void updateCell(int row, int cell, Double v){
        if(row < 17){
            int tableRow = row - 7;
            int tableCell = cell - 1;
            if(row > 13){
                tableRow = tableRow-1;
            }
            DefaultTableModel tableModel = (DefaultTableModel)getTable().getModel();
            //System.out.println(tableRow + "x" + tableCell);
            tableModel.setValueAt(v, tableRow, tableCell);
        }else{
            int tableRow = row - 19;
            int tableCell = cell - 1;
            DefaultTableModel table1Model = (DefaultTableModel)getTable_1().getModel();
          
            Double pv = v*100;
            pv = round(pv, 4);
            table1Model.setValueAt(pv, tableRow, tableCell);
        }
    }
    public static double round(double value, int places) {
        if (places < 0) throw new IllegalArgumentException();

        long factor = (long) Math.pow(10, places);
        value = value * factor;
        long tmp = Math.round(value);
        return (double) tmp / factor;
    }
    
    public WorkbookService getWorkbookService(){
        return workbookService;
    }
}
