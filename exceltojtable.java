package exceltojtable;
import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.*;
import java.util.logging.*;
import javax.swing.table.DefaultTableModel;
import jxl.*;
import jxl.read.biff.BiffException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Workbook;
public class exceltojtable extends JFrame {

    static JTable table;
    static JScrollPane scroll;
    // header is Vector contains table Column
    static Vector headers = new Vector();
    // Model is used to construct JTable
    static DefaultTableModel model = null;
    // data is Vector contains Data from Excel File
    static Vector data = new Vector();
    static JButton jbClick,jbinsert;
    static JFileChooser jChooser;
    static int tableWidth = 0; // set the tableWidth
    static int tableHeight = 0; // set the tableHeight

    public exceltojtable() {
        super("Import Excel To JTable");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        JPanel buttonPanel = new JPanel();
        buttonPanel.setBackground(Color.white);
        jChooser = new JFileChooser();
        jbClick = new JButton("Select Excel File");
        jbinsert = new JButton("Insert Data");
        buttonPanel.add(jbClick, BorderLayout.CENTER);
       
        // Show Button Click Event
        jbClick.addActionListener(new ActionListener() {

            public void actionPerformed(ActionEvent arg0) {
                jChooser.showOpenDialog(null);

                File file = jChooser.getSelectedFile();
                //boolean y=file.getName().endsWith("xlsx");
                
                if (!file.getName().endsWith("xls")&& !file.getName().endsWith("xlsx")) {
                    JOptionPane.showMessageDialog(null,
                            "Please select only Excel file.",
                            "Error", JOptionPane.ERROR_MESSAGE);
                }
                else {
                    try {
                        fillData(file);
                    } catch (IOException ex) {
                        //Logger.getLogger(exceltojtable.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (InvalidFormatException ex) {
                        Logger.getLogger(exceltojtable.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    model = new DefaultTableModel(data,
                            headers);
                    tableWidth = model.getColumnCount()
                            * 150;
                    tableHeight = model.getRowCount()
                            * 25;
                    table.setPreferredSize(new Dimension(
                            tableWidth, tableHeight));

                    table.setModel(model);
                }
            }
        });
         buttonPanel.add(jbinsert, BorderLayout.EAST);
        /*jbinsert.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e) {
                
             new insert().setVisible(true);
            }   
        });*/

        table = new JTable();
        table.setAutoCreateRowSorter(true);
        
        model = new DefaultTableModel(data, headers);

        table.setModel(model);
        table.setBackground(Color.pink);

        table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        table.setEnabled(false);
        table.setRowHeight(25);
        table.setRowMargin(4);

        tableWidth = model.getColumnCount() * 150;
        tableHeight = model.getRowCount() * 25;
        table.setPreferredSize(new Dimension(
                tableWidth, tableHeight));

        scroll = new JScrollPane(table);
        scroll.setBackground(Color.pink);
        scroll.setPreferredSize(new Dimension(300, 300));
        scroll.setHorizontalScrollBarPolicy(
                JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        scroll.setVerticalScrollBarPolicy(
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        getContentPane().add(buttonPanel,
                BorderLayout.NORTH);
        getContentPane().add(scroll,
                BorderLayout.CENTER);
        GraphicsDevice gd = GraphicsEnvironment.getLocalGraphicsEnvironment().getDefaultScreenDevice();
        int width = gd.getDisplayMode().getWidth();
        int height = gd.getDisplayMode().getHeight();
        setSize(width, height);
        setResizable(true);
        setVisible(true);
    }
    
    void fillData(File file) throws IOException, InvalidFormatException {
       Workbook workbook = null;
       workbook=(Workbook) WorkbookFactory.create(file);
       Sheet sheet = workbook.getSheetAt(0);
        int rowStart=sheet.getFirstRowNum();
                    int rowEnd=sheet.getLastRowNum();
       headers.clear();
       for(int i=rowStart;i<=rowStart;i++){
                        Row row=sheet.getRow(i);
       for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
           Cell cell1 = row.getCell(j);
           headers.add(cell1);
       }}
       data.clear();
       for (int k = 1; k <= rowEnd; k++) {
           Vector d = new Vector();
           Row row=sheet.getRow(k);
           for (int l = row.getFirstCellNum(); l < row.getLastCellNum(); l++) {
               Cell cell = row.getCell(l);
               if(cell.getCellType() == cell.CELL_TYPE_NUMERIC){ 
                 int p = (int)cell.getNumericCellValue(); 
                  String strCellValue = String.valueOf(p);
                  d.add(strCellValue);
                        }
                  else { 
                  String strCellValue = cell.toString(); 
                  d.add(strCellValue);
                              }
             
           }
           d.add("\n");
           data.add(d);
       }
                
    }

    public static void main(String[] args) {

        new exceltojtable();
    }
}
