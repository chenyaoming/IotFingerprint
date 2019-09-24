package com.wym.tools;

import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Comparator;
import java.util.List;
import java.util.Vector;
import javax.swing.DefaultRowSorter;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.RowSorter;
import javax.swing.event.TableModelEvent;
import javax.swing.event.TableModelListener;
import javax.swing.filechooser.FileFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
import javax.swing.table.TableRowSorter;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper;

public class ExcelOutandIn extends JFrame implements ActionListener {

    JButton button1 = new JButton("ToExcel");
    JButton button2 = new JButton("FromExcel");

    Container ct = null;
    DefaultTableModel defaultModel = null;

    JButton add = new JButton("添加");
    JButton delete = new JButton("删除");
    JButton save = new JButton("保存");
    JButton reset = new JButton("刷新");

    JPanel jp1 = new JPanel(), jp;
    JPanel jp2 = new JPanel();
    JScrollPane jsp = null;
    UsersDAO userdao = new UsersDAO();
    Users users = null;
    @SuppressWarnings("unchecked")
    List list = null;

    public List getList() {
        return list;
    }

    public void setList(List list) {
        this.list = list;
    }

    protected JTable table = null;
    protected String oldvalue = "";
    protected String newvalue = "";

    /**
     * @param args
     */
    public static void main(String[] args) {
        // TODO Auto-generated method stub

        ExcelOutandIn tm = new ExcelOutandIn();
        // tm.paintuser();
    }

    public void paintuser() {
        list = this.getUsers();
        for (int i = 0; i < list.size(); i++) {
            users = (Users) list.get(i);
            System.out.println(" ID:" + users.getId() + "");
        }
    }

    @SuppressWarnings("unchecked")
    public List getUsers() {
        list = userdao.findAll();

        return list;

    }

    private File getSelectedOpenFile(final String type) {
        String name = getName();

        JFileChooser pathChooser = new JFileChooser();
        pathChooser.setFileFilter(new FileFilter() {

            @Override
            public boolean accept(File f) {
                if (f.isDirectory()) {
                    return true;
                } else {
                    if (f.getName().toLowerCase().endsWith(type)) {
                        return true;
                    } else {
                        return false;
                    }
                }
            }

            @Override
            public String getDescription() {
                return "文件格式（" + type + "）";
            }
        });
        pathChooser.setSelectedFile(new File(name + type));
        int showSaveDialog = pathChooser.showOpenDialog(this);
        if (showSaveDialog == JFileChooser.APPROVE_OPTION) {
            return pathChooser.getSelectedFile();
        } else {
            return null;
        }
    }

    private File getSelectedFile(final String type) {
        String name = getName();

        JFileChooser pathChooser = new JFileChooser();
        pathChooser.setFileFilter(new FileFilter() {

            @Override
            public boolean accept(File f) {
                if (f.isDirectory()) {
                    return true;
                } else {
                    if (f.getName().toLowerCase().endsWith(type)) {
                        return true;
                    } else {
                        return false;
                    }
                }
            }

            @Override
            public String getDescription() {
                return "文件格式（" + type + "）";
            }
        });
        pathChooser.setSelectedFile(new File(name + type));
        int showSaveDialog = pathChooser.showSaveDialog(this);
        if (showSaveDialog == JFileChooser.APPROVE_OPTION) {
            return pathChooser.getSelectedFile();
        } else {
            return null;
        }
    }

    void setlookandfeel() {
        try {

            BeautyEyeLNFHelper.frameBorderStyle = BeautyEyeLNFHelper.FrameBorderStyle.osLookAndFeelDecorated;
            org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper.launchBeautyEyeLNF();
        } catch (final Exception e) {
            System.out.println("error");

        }
    }

    void init() {
        setlookandfeel();

        buildTable();
        jsp = new JScrollPane(table);

        ct.add(jsp);
        button1.setActionCommand("ToExcel");
        button1.addActionListener(this);

        button2.setActionCommand("FromExcel");
        button2.addActionListener(this);

        delete.setActionCommand("delete");
        delete.addActionListener(this);

        reset.setActionCommand("reset");
        reset.addActionListener(this);

        save.setActionCommand("save");
        save.addActionListener(this);

        add.setActionCommand("add");
        add.addActionListener(this);
    }

    public void ToExcel(String path) {

        list = getUsers();

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("Users");

        String[] n = { "编号", "姓名", "密码", "邮箱" };

        Object[][] value = new Object[list.size() + 1][4];
        for (int m = 0; m < n.length; m++) {
            value[0][m] = n[m];
        }
        for (int i = 0; i < list.size(); i++) {
            users = (Users) list.get(i);

            value[i + 1][0] = users.getId();
            value[i + 1][1] = users.getUsername();
            value[i + 1][2] = users.getPassword();
            value[i + 1][3] = users.getUEmail();

        }
        ExcelUtil.writeArrayToExcel(wb, sheet, list.size() + 1, 4, value);

        ExcelUtil.writeWorkbook(wb, path);

    }

    /**
     * 从Excel导入数据到数据库
     * @param filename
     */
    public void FromExcel(String filename) {

        String result = "success";
        /** Excel文件的存放位置。注意是正斜线 */
        // String fileToBeRead = "F:\\" + fileFileName;
        try {
            // 创建对Excel工作簿文件的引用
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(
                    filename));
            // 创建对工作表的引用。
            // HSSFSheet sheet = workbook.getSheet("Sheet1");
            HSSFSheet sheet = workbook.getSheetAt(0);

            int j = 1;//从第2行开始堵数据
            // 第在excel中读取一条数据就将其插入到数据库中
            while (j < sheet.getPhysicalNumberOfRows()) {
                HSSFRow row = sheet.getRow(j);
                Users user = new Users();

                for (int i = 0; i <= 3; i++) {
                    HSSFCell cell = row.getCell((short) i);

                    if (i == 0) {
                        user.setId((int) cell.getNumericCellValue());
                    } else if (i == 1){
                        user.setUsername(cell.getStringCellValue());
                    }

                    else if (i == 2){
                        user.setPassword(cell.getStringCellValue());
                    }


                    else if (i == 3){
                        user.setUEmail(cell.getStringCellValue());
                    }

                }

                System.out.println(user.getId() + " " + user.getUsername()
                        + " " + user.getPassword() + " " + user.getUEmail());

                j++;

                userdao.save(user);

            }

        } catch (FileNotFoundException e2) {
            // TODO Auto-generated catch block
            System.out.println("notfound");
            e2.printStackTrace();
        } catch (IOException e3) {
            // TODO Auto-generated catch block
            System.out.println(e3.toString());

            e3.printStackTrace();
        } catch (Exception e4) {
            System.out.println(e4.toString());

        }

    }

    public JTable CreateTable(String[] columns, Object rows[][]) {
        JTable table;
        TableModel model = new DefaultTableModel(rows, columns);

        table = new JTable(model);
        RowSorter sorter = new TableRowSorter(model);
        table.setRowSorter(sorter);

        return table;

    }

    @SuppressWarnings("unchecked")
    public void fillTable(List<Users> users) {
        DefaultTableModel tableModel = (DefaultTableModel) table.getModel();
        tableModel.setRowCount(0);// 清除原有行

        // 填充数据
        for (Users Users : users) {
            Vector vector = new Vector<Users>();

            vector.add(Users.getId());
            vector.add(Users.getUsername());
            vector.add(Users.getPassword());
            vector.add(Users.getUEmail());

            // 添加数据到表格
            tableModel.addRow(vector);
        }

        // 更新表格
        table.invalidate();
    }

    @SuppressWarnings("unchecked")
    public void tableAddRow(int id, String name, String pwd, String email) {
        DefaultTableModel tableModel = (DefaultTableModel) table.getModel();
        tableModel.getColumnCount();

        // 填充数据

        Vector vector = new Vector<Users>();

        vector.add(id);
        vector.add(name);
        vector.add(pwd);
        vector.add(email);

        // 添加数据到表格
        tableModel.addRow(vector);

        // 更新表格
        table.invalidate();
    }

    @SuppressWarnings("unchecked")
    public void buildTable() {
        String[] n = { "编号", "姓名", "密码", "邮箱" };
        list = getUsers();
        Object[][] value = new Object[list.size()][4];

        for (int i = 0; i < list.size(); i++) {
            users = (Users) list.get(i);

            value[i][0] = users.getId();
            value[i][1] = users.getUsername();
            value[i][2] = users.getPassword();
            value[i][3] = users.getUEmail();

        }
        defaultModel = new DefaultTableModel(value, n) {
            boolean[] editables = { false, true, true, true };

            @Override
            public boolean isCellEditable(int row, int col) {
                return editables[col];
            }
        };
        defaultModel.isCellEditable(1, 1);
        table = new JTable(defaultModel);
        RowSorter sorter = new TableRowSorter(defaultModel);
        table.setRowSorter(sorter);

        // 设置排序
        ((DefaultRowSorter) sorter).setComparator(0, new Comparator<Object>() {
            @Override
            public int compare(Object arg0, Object arg1) {
                try {
                    int a = Integer.parseInt(arg0.toString());
                    int b = Integer.parseInt(arg1.toString());
                    return a - b;
                } catch (NumberFormatException e) {
                    return 0;
                }
            }
        });
        defaultModel.addTableModelListener(new TableModelListener() {

            @Override
            public void tableChanged(TableModelEvent e) {
                if (e.getType() == TableModelEvent.UPDATE) {
                    newvalue = table.getValueAt(e.getLastRow(), e.getColumn())
                            .toString();
                    System.out.println(newvalue);
                    int rowss = table.getEditingRow();
                    if (newvalue.equals(oldvalue)) {
                        System.out.println(rowss);
                        System.out.println(table.getValueAt(rowss, 0) + ""
                                + table.getValueAt(rowss, 1) + ""
                                + table.getValueAt(rowss, 2) + ""
                                + table.getValueAt(rowss, 3));
                        JOptionPane.showMessageDialog(null, "数据没有修改");

                    } else {

                        int dialog = JOptionPane.showConfirmDialog(null,
                                "是否确认修改", "温馨提示", JOptionPane.YES_NO_OPTION);
                        if (dialog == JOptionPane.YES_OPTION) {

                            System.out.println(" 修改了");
                            String s1 = (String) table.getValueAt(rowss, 0)
                                    .toString();
                            int id = Integer.parseInt(s1);
                            users = new Users();
                            users.setId(id);
                            users.setUEmail(table.getValueAt(rowss, 3)
                                    .toString());
                            users.setUsername(table.getValueAt(rowss, 1)
                                    .toString());
                            users.setPassword(table.getValueAt(rowss, 2)
                                    .toString());

                            try {
                                userdao.saveOrUpdate2(users);

                            } catch (Exception eee) {
                                new UsersDAO().saveOrUpdate2(users);
                            }

                        } else if (dialog == JOptionPane.NO_OPTION) {
                            table.setValueAt(oldvalue, rowss, table
                                    .getSelectedColumn());
                            // System.out.println("没有确认修改");
                        }

                    }

                }

            }

        });

        table.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {

                // 记录进入编辑状态前单元格得数据

                try {
                    oldvalue = table.getValueAt(table.getSelectedRow(),
                            table.getSelectedColumn()).toString();
                    System.out.println(oldvalue);
                } catch (Exception ex) {
                    // TODO: handle exception
                }

            }

        });
    }

    public ExcelOutandIn() {

        // TODO Auto-generated constructor stub

        new BorderLayout();

        Font font = new Font("宋体", 4, 14);

        add.setFont(font);
        save.setFont(font);
        delete.setFont(font);
        reset.setFont(font);
        jp1.add(button1);
        jp1.add(button2);

        jp2.add(add);
        jp2.add(delete);
        // jp2.add(save);
        jp2.add(reset);

        ct = this.getContentPane();

        ct.add(jp1, BorderLayout.NORTH);
        ct.add(jp2, BorderLayout.SOUTH);

        init();
        this.setTitle("ToOrFromExcel");
        this.setVisible(true);
        this.setSize(600, 400);
        this.setLocation(400, 250);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

    }

    @Override
    public void actionPerformed(ActionEvent e) {
        // TODO Auto-generated method stub

        if (e.getActionCommand().equals("add")) {

            AddUsers adduser = new AddUsers();
            jp = new JPanel();
            jp.add(adduser);
            ct.add(jp, BorderLayout.WEST);

            /*
             * users= adduser.getU(); if(users==null){
             * JOptionPane.showMessageDialog(null, "Null");
             *
             * }else{
             *
             * }
             */
            // tableAddRow(id, name, pwd, email);
        }
        // defaultModel.addRow(v);
        if (e.getActionCommand().equals("delete")) {
            int srow = 0;

            try {
                srow = table.getSelectedRow();
            } catch (Exception ee) {

            }
            int rowcount = defaultModel.getRowCount() - 1;// getRowCount返回行数，rowcount<0代表已经没有任何行了。

            if (srow > 0) {
                Object id = defaultModel.getValueAt(srow, 0);
                String ID = id.toString();
                users = userdao.findById(Integer.parseInt(ID));

                defaultModel.getRowCount();

                System.out.println(ID);
                defaultModel.removeRow(srow);

                // userdao.delete(users);
                defaultModel.setRowCount(rowcount);
            }
        }

        if (e.getActionCommand().equals("save")) {
            System.out.println("save");
            ct.remove(jp);
        }

        if (e.getActionCommand().equals("reset")) {

            System.out.println("reset");
            fillTable(userdao.findAll());

        }

        if (e.getActionCommand().equalsIgnoreCase("toexcel")) {

            File selectedFile = getSelectedFile(".xls");
            if (selectedFile != null) {
                String path = selectedFile.getPath();

                // System.out.println(path);
                ToExcel(path);
            }

        } else if (e.getActionCommand().equalsIgnoreCase("FromExcel")) {
            File selectedFile = getSelectedOpenFile(".xls");
            if (selectedFile != null) {
                // String name=selectedFile.getName();
                String path = selectedFile.getPath();
                FromExcel(path);
                fillTable(userdao.findAll());

            }

        }

    }

}