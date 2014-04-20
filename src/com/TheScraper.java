package com;

import java.awt.Color;
import java.awt.Toolkit;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.channels.Channels;
import java.nio.channels.ReadableByteChannel;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.ImageIcon;
import javax.swing.SwingWorker;
import javax.swing.UIManager;
import javax.swing.text.BadLocationException;
import javax.swing.text.Style;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

public class TheScraper extends javax.swing.JFrame implements PropertyChangeListener {
    
    Task task;
    FileWriter writer;
    String path;
    Workbook wb = null;
    Sheet sheet = null;
    int r;
    Workbook wb2 = null;
    Sheet sheet2 = null;
    int r2;
    Style style;
    StyledDocument sd;
    
    String cm = "";
    String cs = "";
    int mi = 0;
    int si = 0;
    String startC = "";
    String startS = "";
    int cc;
    boolean commandStop;
    boolean successStoped;
    boolean reachStart = false;
    boolean only = true;
    
    static DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
    
    public TheScraper() {
        initComponents();
        try {
            path = new File(TheScraper.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath()).getParent();
        } catch (URISyntaxException ex) {
            Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, ex);
        }
        System.setProperty("phantomjs.binary.path", path + "\\phantomjs.exe");
        System.setProperty("webdriver.chrome.driver", path + "\\chromedriver.exe");
        this.setIconImage(new ImageIcon(path + "\\lib\\icon.png").getImage());
        AboutD.setIconImage(new ImageIcon(path + "\\lib\\icon.png").getImage());
    }
    
    public String br2nl(String html) {
        Document document = Jsoup.parse(html);
        document.select("br").append("\\n");
        document.select("p").prepend("\\n\\n");
        return document.text().replace("\\n", "\n");
    }
    
    private void createCSVFile() {
        File f = new File(path);
        if (!f.exists()) {
            try {
                writer = new FileWriter(path);
                writer.append("Blog Title");
                writer.append(',');
                writer.append("URL");
                writer.append('\n');
                writer.close();
            } catch (IOException ex) {
                Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
    private void createFile() throws Exception {
        File f = new File(path + "\\available.xls");
        if (f.exists()) {
            readFile();
        } else {
            wb = new HSSFWorkbook();
            sheet = wb.createSheet();
            r = 0;
            Row row = sheet.createRow(r);
            r++;
            row.createCell(0).setCellValue("Title");
            row.createCell(1).setCellValue("Url");
            row.createCell(2).setCellValue("Img");
            //row.createCell(2).setCellValue("Stock");
            row.createCell(3).setCellValue("Price");
            row.createCell(4).setCellValue("Sale Price");
            row.createCell(5).setCellValue("By");
            row.createCell(6).setCellValue("Language");
            row.createCell(7).setCellValue("Format");
            row.createCell(8).setCellValue("Publisher");
            row.createCell(9).setCellValue("Size");
            row.createCell(10).setCellValue("ISPN");
            row.createCell(11).setCellValue("Topic");
            row.createCell(12).setCellValue("Details");
        }
    }
    
    private void createFile2() throws Exception {
        File f = new File(path + "\\out of stock.xls");
        if (f.exists()) {
            readFile2();
        } else {
            wb2 = new HSSFWorkbook();
            sheet2 = wb2.createSheet();
            r2 = 0;
            Row row = sheet2.createRow(r2);
            r2++;
            row.createCell(0).setCellValue("Title");
            row.createCell(1).setCellValue("Url");
            row.createCell(2).setCellValue("Img");
            row.createCell(3).setCellValue("Price");
            row.createCell(4).setCellValue("Sale Price");
            row.createCell(5).setCellValue("By");
            row.createCell(6).setCellValue("Language");
            row.createCell(7).setCellValue("Format");
            row.createCell(8).setCellValue("Publisher");
            row.createCell(9).setCellValue("Size");
            row.createCell(10).setCellValue("ISPN");
            row.createCell(11).setCellValue("Topic");
            row.createCell(12).setCellValue("Details");
        }
    }
    
    private void readFile() {
        try {
            FileInputStream file = new FileInputStream(new File(path + "\\available.xls"));
            wb = new HSSFWorkbook(file);
            sheet = wb.getSheetAt(0);
            r = sheet.getLastRowNum() + 1;
        } catch (Exception e) {
            Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, e);
        }
    }
    
    private void readFile2() {
        try {
            FileInputStream file = new FileInputStream(new File(path + "\\out of stock.xls"));
            wb2 = new HSSFWorkbook(file);
            sheet2 = wb2.getSheetAt(0);
            r2 = sheet2.getLastRowNum() + 1;
        } catch (Exception e) {
            Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, e);
        }
    }
    
    private void closeFile() {
        try {
            if (wb != null) {
                int i;
                try {
                    for (i = 0; i < 12; i++) {
                        sheet.autoSizeColumn(i);
                    }
                } catch (Exception e) {
                }
                
                FileOutputStream fileOut = new FileOutputStream(path + "\\available.xls");
                wb.write(fileOut);
                fileOut.close();
                wb = null;
            }
        } catch (Exception e) {
            Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, e);
            errorLog("error saving the spreadsheet");
        }
    }
    
    private void closeFile2() {
        try {
            if (wb2 != null) {
                int i;
                try {
                    for (i = 0; i < 12; i++) {
                        sheet2.autoSizeColumn(i);
                    }
                } catch (Exception e) {
                }
                
                FileOutputStream fileOut = new FileOutputStream(path + "\\out of stock.xls");
                wb2.write(fileOut);
                fileOut.close();
                wb2 = null;
            }
        } catch (Exception e) {
            Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, e);
            errorLog("error saving the spreadsheet");
        }
    }
    
    private void normalLog(String s) {
        try {
            StyleConstants.setForeground(style, Color.black);
            sd.insertString(sd.getLength(), s + "\n", style);
        } catch (BadLocationException e) {
        }
    }
    
    private void errorLog(String s) {
        try {
            StyleConstants.setForeground(style, Color.red);
            sd.insertString(sd.getLength(), s + "\n", style);
        } catch (BadLocationException e) {
        }
    }
    
    private void successLog(String s) {
        try {
            StyleConstants.setForeground(style, new Color(39245));
            sd.insertString(sd.getLength(), s + "\n", style);
        } catch (BadLocationException e) {
        }
    }
    
    private String getParam(Map<String, String> hm, String s) {
        String x = "";
        if (hm.containsKey(s)) {
            x = String.valueOf(hm.get(s));
        }
        return x;
    }
    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        AboutD = new javax.swing.JDialog();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        jLabel11 = new javax.swing.JLabel();
        imgL = new javax.swing.JLabel();
        executeBut = new javax.swing.JButton();
        progressBar = new javax.swing.JProgressBar();
        display = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        log = new javax.swing.JTextPane();
        jLabel1 = new javax.swing.JLabel();
        catCombo = new javax.swing.JComboBox();
        subCombo = new javax.swing.JComboBox();
        jLabel2 = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenuItem1 = new javax.swing.JMenuItem();
        jMenu2 = new javax.swing.JMenu();
        jMenuItem2 = new javax.swing.JMenuItem();

        AboutD.setTitle("About Me");
        AboutD.addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowOpened(java.awt.event.WindowEvent evt) {
                AboutDWindowOpened(evt);
            }
        });

        jLabel3.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        jLabel3.setText("My name");

        jLabel4.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        jLabel4.setText("Riadh Chebbi");

        jLabel5.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        jLabel5.setText("Odesk");

        jLabel6.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        jLabel6.setText("riadh-c");

        jLabel7.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        jLabel7.setText("riadh_chebbi");

        jLabel8.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        jLabel8.setText("Skype");

        jLabel9.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        jLabel9.setText("Email");

        jLabel10.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        jLabel10.setText("riadh_chebbi@hotmail.fr");

        jLabel11.setFont(new java.awt.Font("Verdana", 0, 14)); // NOI18N
        jLabel11.setForeground(new java.awt.Color(255, 0, 51));
        jLabel11.setText("God of Web Scraping at your service");

        javax.swing.GroupLayout AboutDLayout = new javax.swing.GroupLayout(AboutD.getContentPane());
        AboutD.getContentPane().setLayout(AboutDLayout);
        AboutDLayout.setHorizontalGroup(
            AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(AboutDLayout.createSequentialGroup()
                .addGap(36, 36, 36)
                .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel11, javax.swing.GroupLayout.PREFERRED_SIZE, 310, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(AboutDLayout.createSequentialGroup()
                        .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                .addComponent(jLabel8, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jLabel3, javax.swing.GroupLayout.DEFAULT_SIZE, 81, Short.MAX_VALUE)
                                .addComponent(jLabel5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                            .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 103, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(24, 24, 24)
                        .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 101, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 92, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 92, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 174, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(imgL, javax.swing.GroupLayout.PREFERRED_SIZE, 203, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(46, Short.MAX_VALUE))
        );
        AboutDLayout.setVerticalGroup(
            AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(AboutDLayout.createSequentialGroup()
                .addGap(20, 20, 20)
                .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(imgL, javax.swing.GroupLayout.PREFERRED_SIZE, 207, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(AboutDLayout.createSequentialGroup()
                        .addComponent(jLabel11, javax.swing.GroupLayout.PREFERRED_SIZE, 34, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(18, 18, 18)
                        .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 28, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel4))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel5, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel7, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(AboutDLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel9, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addContainerGap(28, Short.MAX_VALUE))
        );

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("YouTubeTranscription 1.00");
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowClosing(java.awt.event.WindowEvent evt) {
                formWindowClosing(evt);
            }
            public void windowOpened(java.awt.event.WindowEvent evt) {
                formWindowOpened(evt);
            }
        });

        executeBut.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        executeBut.setText("execute");
        executeBut.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                executeButActionPerformed(evt);
            }
        });

        progressBar.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        progressBar.setStringPainted(true);

        display.setFont(new java.awt.Font("Verdana", 0, 14)); // NOI18N
        display.setForeground(new java.awt.Color(51, 51, 255));

        log.setEditable(false);
        log.setBackground(javax.swing.UIManager.getDefaults().getColor("Button.background"));
        log.setBorder(javax.swing.BorderFactory.createTitledBorder(""));
        log.setFont(new java.awt.Font("Verdana", 0, 13)); // NOI18N
        log.setMargin(new java.awt.Insets(17, 17, 3, 3));
        jScrollPane1.setViewportView(log);

        jLabel1.setFont(new java.awt.Font("Verdana", 1, 13)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(51, 51, 255));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("Log");
        jLabel1.setBorder(javax.swing.BorderFactory.createTitledBorder(""));

        catCombo.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                catComboActionPerformed(evt);
            }
        });

        jLabel2.setFont(new java.awt.Font("Verdana", 1, 11)); // NOI18N
        jLabel2.setText("Category :");

        jLabel12.setFont(new java.awt.Font("Verdana", 1, 11)); // NOI18N
        jLabel12.setText("sub-Category :");

        jMenu1.setText("File");
        jMenu1.setFont(new java.awt.Font("Verdana", 0, 15)); // NOI18N

        jMenuItem1.setText("Quit");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem1);

        jMenuBar1.add(jMenu1);

        jMenu2.setText("About");
        jMenu2.setFont(new java.awt.Font("Verdana", 0, 15)); // NOI18N

        jMenuItem2.setText("About me");
        jMenuItem2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem2ActionPerformed(evt);
            }
        });
        jMenu2.add(jMenuItem2);

        jMenuBar1.add(jMenu2);

        setJMenuBar(jMenuBar1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(25, 25, 25)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(progressBar, javax.swing.GroupLayout.DEFAULT_SIZE, 452, Short.MAX_VALUE)
                            .addComponent(jScrollPane1))
                        .addGap(18, 18, 18)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel2)
                            .addComponent(jLabel12))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(subCombo, javax.swing.GroupLayout.PREFERRED_SIZE, 287, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(catCombo, javax.swing.GroupLayout.PREFERRED_SIZE, 287, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(executeBut, javax.swing.GroupLayout.PREFERRED_SIZE, 158, javax.swing.GroupLayout.PREFERRED_SIZE)))
                    .addComponent(display, javax.swing.GroupLayout.PREFERRED_SIZE, 647, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 51, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(82, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(21, 21, 21)
                .addComponent(display, javax.swing.GroupLayout.PREFERRED_SIZE, 45, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 28, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 364, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(catCombo, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel2))
                        .addGap(31, 31, 31)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(subCombo, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel12))
                        .addGap(224, 224, 224)
                        .addComponent(executeBut)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 7, Short.MAX_VALUE)
                .addComponent(progressBar, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(28, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    class Task extends SwingWorker<String, String> {
        
        @Override
        public String doInBackground() {
            setProgress(0);
            try {
                act();
            } catch (Exception ex) {
                Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, ex);
                errorLog(ex.toString());
            }
            return null;
        }
        
        private void act() throws Exception {
            createFile();
            createFile2();
            for (int i = catCombo.getSelectedIndex(); i < catCombo.getItemCount(); i++) {
                cm = ((Category) catCombo.getItemAt(i)).getS();
                normalLog("--- " + cm + " ---");
                (new File(path + "\\" + cm)).mkdir();
                for (int j = subCombo.getSelectedIndex(); j < subCombo.getItemCount(); j++) {
                    cs = ((SubCategory) subCombo.getItemAt(j)).getS();
                    normalLog("+ " + cs);
                    (new File(path + "\\" + cm + "\\" + cs)).mkdir();
                    String url = "http://www.alkitab.com/" + ((SubCategory) subCombo.getItemAt(j)).getUrl();
                    getPage(url);
                    return;
                }
                try {
                    catCombo.setSelectedIndex(i + 1);
                } catch (Exception e) {
                }
            }
        }
        
        void getPage(String url) throws Exception {
            //successLog(url);
            Document doc = getHtm(url);
            if (doc == null) {
                return;
            }
            try {
                Elements li = doc.getElementById("div1").getElementsByTag("li");
                for (int i = 0; i < li.size(); i++) {
                    String u = "http://www.alkitab.com/" + li.get(i).getElementsByTag("a").first().attr("href");
                    publish("scraping record " + (i + 1) + " / " + li.size());
                    setProgress((int) ((i + 1) * 100 / li.size()));
                    try {
                        getRecord(u);
                    } catch (Exception e) {
                        Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, u);
                    }
                    
                }
            } catch (Exception e) {
                errorLog("err ");
            }
        }
        
        void getRecord(String url) throws Exception {
            Document doc1 = Jsoup.connect(url).timeout(120000).get();
            String title = "";
            try {
                title = doc1.getElementById("item-contenttitle").text().replace(",", ".");
            } catch (Exception e) {
                title = doc1.getElementById("section-contenttitle").text().replace(",", ".");
            }
            String av = "not available on stock";
            try {
                Elements r = doc1.getElementsByClass("addtocartImg");
                if (r.size() == 1) {
                    av = "available on stock";
                }
            } catch (Exception e) {
            }
            /*  if (av.equals("available on stock")) {
             return;
             }*/
            String price = "";
            String salePrice = "";
            try {
                price = doc1.getElementsByClass("price-bold").first().text();
            } catch (Exception e) {
                price = doc1.getElementsByClass("price").first().text().replace("Price:", "");
            }
            try {
                salePrice = doc1.getElementsByClass("sale-price-bold").first().getElementsByTag("em").first().text();
            } catch (Exception e) {
            }
            String det = "";
            String more = "";
            Map hm = new HashMap<String, String>();
            int i;
            try {
                det = (doc1.getElementById("contents").text()).replace(",", ".");
            } catch (Exception x) {
                return;
            }
            more = det.substring(det.indexOf("More About This Item") + 20);
            det = det.substring(0, det.indexOf("More About This Item"));
            String d[] = det.split(":");
            String s1[] = new String[d.length];
            String s2[] = new String[d.length];
            
            s1[0] = d[0];
            s1[0] = s1[0].replace("Details ", "");
            s2[0] = d[1];
            for (i = 1; i < d.length - 1; i++) {
                s1[i] = d[i].substring(d[i].lastIndexOf(" ") + 1);
                s2[i] = d[i + 1];
                //System.out.println(s2[i]);
            }
            for (i = 0; i < s1.length - 2; i++) {
                s2[i] = s2[i].replace(s1[i + 1], "");
            }
            
            hm.put("More", more);
            for (i = 0; i < s1.length; i++) {
                // System.out.println(s1[i] + " : " + s2[i]);
                hm.put(s1[i], s2[i]);
            }
            String u = "";
            try {
                u = doc1.getElementById("itemarea").getElementsByTag("img").first().attr("src");
            } catch (Exception e) {
                u = doc1.getElementById("caption").getElementsByTag("img").first().attr("src");
            }
            FileOutputStream fos = null;
            String sse = "";
            
            try {
                URL iurl = new URL(u);
                ReadableByteChannel rbc = Channels.newChannel(iurl.openStream());
                if (sse.length() > 200) {
                    sse = sse.substring(0, 200);
                }
                //  System.out.println(title);
                sse = title.replace("\\", "").replaceAll("\\W", "");
                sse = sse.replace("/", "");
                sse = sse.replaceAll("[\\/?*:'<>|\"]", "");
                fos = new FileOutputStream(path + "\\" + cm + "\\" + cs + "\\" + sse + ".jpg");
                fos.getChannel().transferFrom(rbc, 0, Long.MAX_VALUE);
                fos.flush();
                fos.close();
            } catch (Exception ex) {
                System.out.println(sse);
                Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, ex);
            }
            if (av.equals("available on stock")) {
                Row row = sheet.createRow(r);
                r++;
                // writer = new FileWriter(path + "\\" + cm + "\\" + cs + "\\" + cs + ".csv");
                row.createCell(0).setCellValue(title);
                row.createCell(1).setCellValue(url);
                row.createCell(2).setCellValue(sse + ".jpg");
                // row.createCell(2).setCellValue(av);
                row.createCell(3).setCellValue(price);
                row.createCell(4).setCellValue(salePrice);
                row.createCell(5).setCellValue(getParam(hm, "By"));
                row.createCell(6).setCellValue(getParam(hm, "Language"));
                row.createCell(7).setCellValue(getParam(hm, "Format"));
                row.createCell(8).setCellValue(getParam(hm, "Publisher"));
                row.createCell(9).setCellValue(getParam(hm, "Size"));
                row.createCell(10).setCellValue(getParam(hm, "ISBN"));
                row.createCell(11).setCellValue(getParam(hm, "Topic"));
                row.createCell(12).setCellValue(getParam(hm, "More"));
            } else {
                Row row = sheet2.createRow(r2);
                r2++;
                // writer = new FileWriter(path + "\\" + cm + "\\" + cs + "\\" + cs + ".csv");
                row.createCell(0).setCellValue(title);
                row.createCell(1).setCellValue(url);
                row.createCell(2).setCellValue(sse + ".jpg");
                // row.createCell(2).setCellValue(av);
                row.createCell(3).setCellValue(price);
                row.createCell(4).setCellValue(salePrice);
                row.createCell(5).setCellValue(getParam(hm, "By"));
                row.createCell(6).setCellValue(getParam(hm, "Language"));
                row.createCell(7).setCellValue(getParam(hm, "Format"));
                row.createCell(8).setCellValue(getParam(hm, "Publisher"));
                row.createCell(9).setCellValue(getParam(hm, "Size"));
                row.createCell(10).setCellValue(getParam(hm, "ISBN"));
                row.createCell(11).setCellValue(getParam(hm, "Topic"));
                row.createCell(12).setCellValue(getParam(hm, "More"));
            }
        }
        
        Document getHtm(String url) throws Exception {
            int x = 0;
            Document doc = null;
            do {
                try {
                    doc = Jsoup.connect(url).timeout(60000).get();
                    return doc;
                } catch (Exception e) {
                    errorLog("connection error");
                    if (x == 3) {
                        return null;
                    }
                    Thread.sleep(5000);
                    x++;
                }
            } while (true);
        }
        
        @Override
        protected void process(List<String> chunks) {
            display.setText(chunks.get(0));
        }
        
        @Override
        public void done() {
            Toolkit.getDefaultToolkit().beep();
            executeBut.setEnabled(true);
            setCursor(null); //turn off the wait cursor
        }
    }
    
    @Override
    public void propertyChange(PropertyChangeEvent evt) {
        if ("progress".equals(evt.getPropertyName())) {
            int progress = (Integer) evt.getNewValue();
            progressBar.setValue(progress);
            
        }
    }
    
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            // UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            UIManager.setLookAndFeel("com.seaglasslookandfeel.SeaGlassLookAndFeel");
        } catch (Exception e) {
            e.printStackTrace();
        }
        /* try {
         for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
         if ("Nimbus".equals(info.getName())) {
         javax.swing.UIManager.setLookAndFeel(info.getClassName());
         break;
         }
         }
         } catch (ClassNotFoundException ex) {
         java.util.logging.Logger.getLogger(Scraper.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
         } catch (InstantiationException ex) {
         java.util.logging.Logger.getLogger(Scraper.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
         } catch (IllegalAccessException ex) {
         java.util.logging.Logger.getLogger(Scraper.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
         } catch (javax.swing.UnsupportedLookAndFeelException ex) {
         java.util.logging.Logger.getLogger(Scraper.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
         }
         //</editor-fold>

         /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new TheScraper().setVisible(true);
            }
        });
    }

    private void executeButActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_executeButActionPerformed
        //setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        executeBut.setEnabled(false);
        task = new Task();
        task.addPropertyChangeListener(this);
        
        mi = 0;
        si = 0;
        cm = "";
        cs = "";
        cc = 0;
        commandStop = false;
        reachStart = false;
        //only = onlyC.isSelected();
        task.execute();
    }//GEN-LAST:event_executeButActionPerformed

    private void jMenuItem2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem2ActionPerformed
        AboutD.setSize(600, 300);
        AboutD.setLocationRelativeTo(this);
        AboutD.setVisible(true);
    }//GEN-LAST:event_jMenuItem2ActionPerformed

    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        System.exit(0);
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    private void AboutDWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_AboutDWindowOpened
        imgL.setIcon(new ImageIcon(path + "\\lib\\riadh.jpg"));
    }//GEN-LAST:event_AboutDWindowOpened

    private void formWindowClosing(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowClosing
        closeFile();
        closeFile2();
    }//GEN-LAST:event_formWindowClosing

    private void formWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowOpened
        style = log.addStyle("I'm a Style", null);
        sd = log.getStyledDocument();
        jScrollPane1.setBorder(null);
        
        Document doc1 = null;
        normalLog("fetching categories from the website");
        try {
            doc1 = Jsoup.connect("http://www.alkitab.com/").timeout(120000).get();
        } catch (IOException ex) {
            errorLog("Error connecting to the website");
            Logger.getLogger(TheScraper.class.getName()).log(Level.SEVERE, null, ex);
            return;
        }
        Elements menus = doc1.getElementById("nav-menu").getElementsByTag("li");
        Elements cate = doc1.getElementById("nav-menu").getElementsByTag("ul").first().children();
        List<Category> lc = new ArrayList<>();
        for (int i = 0; i < cate.size(); i++) {
            List<SubCategory> ls = new ArrayList<>();
            Elements sub = cate.get(i).getElementsByClass("f2");
            for (int j = 0; j < sub.size(); j++) {
                String title = sub.get(j).getElementsByTag("a").text().replaceAll("[\\/?*:'<>|]", "");
                String href = sub.get(j).getElementsByTag("a").attr("href");
                ls.add(new SubCategory(title, href));
            }
            String catTtitle = cate.get(i).getElementsByTag("a").first().text().replaceAll("[\\/?*:'<>|]", "");

            //lc.add(new Category(catTtitle, ls));
            catCombo.addItem(new Category(catTtitle, ls));
        }
        successLog("categories populated");
    }//GEN-LAST:event_formWindowOpened

    private void catComboActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_catComboActionPerformed
        Category c = (Category) catCombo.getSelectedItem();
        int i;
        subCombo.removeAllItems();
        //subCombo.addItem("ALL");
        for (i = 0; i < c.getLs().size(); i++) {
            subCombo.addItem(c.getLs().get(i));
        }
    }//GEN-LAST:event_catComboActionPerformed

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JDialog AboutD;
    private javax.swing.JComboBox catCombo;
    private javax.swing.JLabel display;
    private javax.swing.JButton executeBut;
    private javax.swing.JLabel imgL;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItem2;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTextPane log;
    private javax.swing.JProgressBar progressBar;
    private javax.swing.JComboBox subCombo;
    // End of variables declaration//GEN-END:variables
}
