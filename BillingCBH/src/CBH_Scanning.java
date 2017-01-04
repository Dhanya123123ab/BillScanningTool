import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.StringTokenizer;

import javax.swing.JPasswordField;

import jxl.Cell;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.Number;
import jxl.write.Label;
import jxl.write.Boolean;
import jxl.write.DateTime;

import org.apache.poi.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class CBH_Scanning extends JFrame {

	private JPanel contentPane;
	private JTextField LogInID;
	private JTextField Sheet;
	private JPasswordField password;
	String x,y,z,result;
	String b[]=null;
	String a[]=null;

	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					CBH_Scanning frame = new CBH_Scanning();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public CBH_Scanning() {
		setTitle("Scanning");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 871, 737);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JLabel lblCadmLoginId = new JLabel("CADM Login ID : ");
		lblCadmLoginId.setBounds(87, 83, 109, 33);
		contentPane.add(lblCadmLoginId);
		
		LogInID = new JTextField();
		LogInID.setBounds(219, 83, 164, 33);
		contentPane.add(LogInID);
		LogInID.setColumns(10);
		
		JLabel lblCadmDbPassword = new JLabel("CADM DB Password :");
		lblCadmDbPassword.setBounds(87, 168, 127, 28);
		contentPane.add(lblCadmDbPassword);
		
		JLabel lblEnv = new JLabel("Env. : ");
		lblEnv.setBounds(87, 259, 72, 16);
		contentPane.add(lblEnv);
		
		JComboBox Env = new JComboBox();
		Env.setModel(new DefaultComboBoxModel(new String[] {"Select ....", "UAT", "ST", "ILT"}));
		Env.setBounds(219, 253, 164, 28);
		contentPane.add(Env);
		
		JLabel lblCbhListSheet = new JLabel("CBH List Sheet : ");
		lblCbhListSheet.setBounds(87, 337, 127, 28);
		contentPane.add(lblCbhListSheet);
		
		Sheet = new JTextField();
		Sheet.setBounds(219, 337, 252, 28);
		contentPane.add(Sheet);
		Sheet.setColumns(10);
		
		JButton btnBrowse = new JButton("Browse");
		btnBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (arg0.getSource() == btnBrowse)
			    {
			        JFileChooser chooser = new JFileChooser(new File(System.getProperty("user.home") + "\\Libraries")); //Libraries Directory as default
			        chooser.setDialogTitle("Select Location");
			        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
			        chooser.setAcceptAllFileFilterUsed(false);

			        if (chooser.showSaveDialog(chooser) == JFileChooser.APPROVE_OPTION)
			        { 
			            String fileID = chooser.getSelectedFile().getPath();
			            Sheet.setText(fileID);
			            
			        }
			    }
			}
		});
		btnBrowse.setBounds(519, 337, 109, 28);
		contentPane.add(btnBrowse);
		
		JButton btnRun = new JButton("Run");
		btnRun.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				int selection=Env.getSelectedIndex();
				switch(selection)
				{
				case 0: JOptionPane.showMessageDialog(null, "Please select Environment");
				break;
				case 1: 
					
					y=LogInID.getText();
					z=password.getText();
					x=Sheet.getText();
					
					try
					{
						HSSFWorkbook wb=new HSSFWorkbook();
						 HSSFSheet sh1=wb.createSheet("Result");
						 HSSFRow myrow=sh1.createRow(0);
							HSSFCell mycol=myrow.createCell(0);
							mycol.setCellValue("CBH");
							HSSFCell mycol1=myrow.createCell(1);
							mycol1.setCellValue("Result");
						 File f=new File(x);
						Workbook w=Workbook.getWorkbook(f);
						jxl.Sheet sh=w.getSheet(0);
						int row=sh.getRows();
						int columns=sh.getColumns();
						for(int k=0;k<row;k++)
						{
							for(int g=0;g<columns;g++)
							{
								Cell c=sh.getCell(g,k);
								
								result+=c.getContents()+"\",\"";
								
							}
						}
						 b=result.split("\",\"");
						 //a=b[0].split(",");
						 StringTokenizer tok = new StringTokenizer(b[0],",;");	
						 //System.out.println(tok.countTokens());
						 int k=tok.countTokens();
						 a = new String[k];
						for(int i=0;i<k;i++)
						{
							a[i]=tok.nextToken();
							System.out.println(a[i]);
						}
						
						int j=a.length;
						    for(int i=0;i<j;i++){
						    	
						    	
						    	if(a[i].contains("#"))
						    	{
						    		//c[i]=a[i].substring(0,a[i].lastIndexOf('#')-1);
						    		a[i]=a[i].substring(a[i].lastIndexOf('#')+1);
						    		
						    	}
						    	System.out.println(a[i]);
						    }
						Connection con = DriverManager.getConnection("jdbc:oracle:thin:@hltv0754.hydc.sbc.com:1524:CADMU1DB",y,z);
						Statement stmt=con.createStatement(); 
						Statement stmt1=con.createStatement();
						Statement stmt2=con.createStatement();
						Statement stmt3=con.createStatement();
						for(int i=0;i<a.length;i++)
						{
							
						ResultSet rs=stmt.executeQuery("select distinct hier.cust_blng_hier_id,hier.acct_1_nb,bus.rpt_cyc_cd,bus.ub_cyc_cd,cust.blk_extrct_nd,hier.cfm_strt_dt,hier.ub_acct_type_cd,CNTRY_CD,cust.MNL_BL_ND,cust.PARTL_HIER_ND, hier.VRTL_INV_ND,bus.BU_STAT_CD age from hier_pnt_tb hier,cust_blng_hier_tb cust,bus_arng_tb bus where hier.cust_blng_hier_id = cust.cust_blng_hier_id and bus.cust_blng_hier_id = cust.cust_blng_hier_id and bus.hier_pnt_id = hier.hier_pnt_id and hier.cust_blng_hier_id in ("+a[i]+") and hier.ub_acct_type_cd=bus.bus_arng_type_cd"); 
						ResultSet rs1=stmt1.executeQuery("Select acct_1_nb,cust_blng_hier_id,cfm_strt_dt,VAT_RGSTRN_NB,VAT_TAX_CD,VAT_EXMPT_CRTFCATE_NB,row_creat_ts, row_updt_ts,CNTRY_CD,rgnl_blr from hier_pnt_tb hier where cust_blng_hier_id in("+a[i]+") and ub_acct_type_cd in ('C')");
						ResultSet rs2=stmt2.executeQuery("select acct_1_nb,cust_blng_hier_id,cfm_strt_dt,VAT_TAX_CD, HIER_PNT_DESC_UTF8_TX,VAT_EXMPT_CRTFCATE_NB,CNTRY_CD,rgnl_blr from hier_pnt_tb hier where cust_blng_hier_id in("+a[i]+") and ub_acct_type_cd in ('C','I','AB','AG','AS')");
						ResultSet rs3=stmt3.executeQuery("select cbh.cust_blng_hier_id, cust.MSTR_AGMT_NB, cust.CUST_ACSS_AUTH_CD from CUST_TB cust, CUST_BLNG_HIER_TB cbh where cbh.cust_id = cust.cust_id and cust_blng_hier_id in ("+a[i]+")");
						while(rs.next()&&rs1.next()&&rs2.next()&&rs3.next()){
							HSSFRow myr=sh1.createRow(i+1);
							HSSFCell myc=myr.createCell(0);
							myc.setCellValue(rs.getString(1));
							if(rs.getString(10)==null)//Partial Indicator= NULL
							{
								if(rs.getString(9)==null)//Manual Indicator= NULL
								{
									if(rs.getString(5)==null || rs.getString(5)=="Y")//Block Extract Indicator= ‘Y’ or NULL
									{
										if(rs.getString(6)!=null)//CFM start Date !=null
										{
											if(rs.getInt(12)<5)//Age Control shouldn’t be more than5
											{
												if(rs1.getString(3)!=null)//CFM start Date !=null in Query 2
												{
													if(rs2.getString(5)==null)
													{
														if(rs3.getString(2)==null && rs1.getString(9)=="UB-MOW")
														{
																																					
														HSSFCell myc1=myr.createCell(1);
														myc1.setCellValue("MA# is Mandatory for "+rs1.getString(9));
														}
														else
														{
															HSSFCell myc1=myr.createCell(1);
															myc1.setCellValue("Valid");
														}
													}
													else
													{
													HSSFCell myc1=myr.createCell(1);
													myc1.setCellValue("HIER_PNT_DESC_UTF8_TX value is "+ rs2.getString(5));
														
													}
												}
												else
												{
													HSSFCell myc1=myr.createCell(1);
													myc1.setCellValue("CFM Start date is present in Query 1 but not Present in Query 2 for "+ rs1.getString(9) +" country(US or UB-MOW Mandatory)");
													
												}
											}
											else
											{
												HSSFCell myc1=myr.createCell(1);
												myc1.setCellValue("Age control is "+rs.getString(12));
											}
										
										}
										else
										{
											HSSFCell myc1=myr.createCell(1);
											//myc1.setCellValue("CFM Start date is not present in Query 1");
											if(rs1.getString(3)==null)
											{
												myc1.setCellValue("CFM Start date is not present in Query 1 and Query 2 for "+ rs1.getString(9) +"country(US or UB-MOW Mandatory)");
											}
											
											else
											{
												myc1.setCellValue("CFM Start Date is not present in Query 1 but present in Query 2 for "+ rs1.getString(9) +"country(US or UB-MOW Mandatory)");
											}
											
										}
									}
									else
									{
										HSSFCell myc1=myr.createCell(1);
										myc1.setCellValue("Block Indicator is "+rs.getString(5));
									}
								
								}
								else
								{
									HSSFCell myc1=myr.createCell(1);
									myc1.setCellValue("Manual Indicator is "+rs.getString(9));
								}
									
							}
							else
							{
								HSSFCell myc1=myr.createCell(1);
								myc1.setCellValue("Partial Indicator is "+rs.getString(10));	
							}
							
							
						}
						FileOutputStream f1=new FileOutputStream(new File("C:\\ScannedData.xls"));
						wb.write(f1);
						
						f1.close();
						rs.close();
						rs1.close();
						rs2.close();
						rs3.close();
						
						}
						stmt.close();
						stmt1.close();
						stmt2.close();
						stmt3.close();
						
						con.close();
						a=null;
						b=null;
						
					}
					catch(Exception e1)
					{
						e1.printStackTrace();
					}
					JOptionPane.showMessageDialog(null, "Done");
					
				break;
				case 2:
					
					/*y=LogInID.getText();
					z=password.getText();
					x=Sheet.getText();
					
					try
					{
						HSSFWorkbook wb=new HSSFWorkbook();
						 HSSFSheet sh1=wb.createSheet("Result");
						 HSSFRow myrow=sh1.createRow(0);
							HSSFCell mycol=myrow.createCell(0);
							mycol.setCellValue("CBH");
							HSSFCell mycol1=myrow.createCell(1);
							mycol1.setCellValue("Result");
						 File f=new File(x);
						Workbook w=Workbook.getWorkbook(f);
						jxl.Sheet sh=w.getSheet(0);
						int row=sh.getRows();
						int columns=sh.getColumns();
						for(int k=0;k<row;k++)
						{
							for(int g=0;g<columns;g++)
							{
								Cell c=sh.getCell(g,k);
								
								result+=c.getContents()+"\",\"";
								
							}
						}
						 b=result.split("\",\"");
						 //a=b[0].split(",");
						 StringTokenizer tok = new StringTokenizer(b[0],",;");	
						 //System.out.println(tok.countTokens());
						 int k=tok.countTokens();
						 a = new String[k];
						for(int i=0;i<k;i++)
						{
							a[i]=tok.nextToken();
							//System.out.println(a[i]);
						}
						
						int j=a.length;
						    for(int i=0;i<j;i++){
						    	
						    	
						    	if(a[i].contains("#"))
						    	{
						    		//c[i]=a[i].substring(0,a[i].lastIndexOf('#')-1);
						    		a[i]=a[i].substring(a[i].lastIndexOf('#')+1);
						    		
						    	}
						    	System.out.println(a[i]);
						    }
						Connection con = DriverManager.getConnection("jdbc:oracle:thin:@hltv0754.hydc.sbc.com:1524:CADMU1DB",y,z);
						Statement stmt=con.createStatement(); 
						Statement stmt1=con.createStatement();
						Statement stmt2=con.createStatement();
						Statement stmt3=con.createStatement();
						for(int i=0;i<a.length;i++)
						{
							
						ResultSet rs=stmt.executeQuery("select distinct hier.cust_blng_hier_id,hier.acct_1_nb,bus.rpt_cyc_cd,bus.ub_cyc_cd,cust.blk_extrct_nd,hier.cfm_strt_dt,hier.ub_acct_type_cd,CNTRY_CD,cust.MNL_BL_ND,cust.PARTL_HIER_ND, hier.VRTL_INV_ND,bus.BU_STAT_CD age from hier_pnt_tb hier,cust_blng_hier_tb cust,bus_arng_tb bus where hier.cust_blng_hier_id = cust.cust_blng_hier_id and bus.cust_blng_hier_id = cust.cust_blng_hier_id and bus.hier_pnt_id = hier.hier_pnt_id and hier.cust_blng_hier_id in ("+a[i]+") and hier.ub_acct_type_cd=bus.bus_arng_type_cd"); 
						ResultSet rs1=stmt1.executeQuery("Select acct_1_nb,cust_blng_hier_id,cfm_strt_dt,VAT_RGSTRN_NB,VAT_TAX_CD,VAT_EXMPT_CRTFCATE_NB,row_creat_ts, row_updt_ts,CNTRY_CD,rgnl_blr from hier_pnt_tb hier where cust_blng_hier_id in("+a[i]+") and ub_acct_type_cd in ('C')");
						ResultSet rs2=stmt2.executeQuery("select acct_1_nb,cust_blng_hier_id,cfm_strt_dt,VAT_TAX_CD, HIER_PNT_DESC_UTF8_TX,VAT_EXMPT_CRTFCATE_NB,CNTRY_CD,rgnl_blr from hier_pnt_tb hier where cust_blng_hier_id in("+a[i]+") and ub_acct_type_cd in ('C','I','AB','AG','AS')");
						ResultSet rs3=stmt3.executeQuery("select cbh.cust_blng_hier_id, cust.MSTR_AGMT_NB, cust.CUST_ACSS_AUTH_CD from CUST_TB cust, CUST_BLNG_HIER_TB cbh where cbh.cust_id = cust.cust_id and cust_blng_hier_id in ("+a[i]+")");
						while(rs.next()&&rs1.next()&&rs2.next()&&rs3.next()){
							HSSFRow myr=sh1.createRow(i+1);
							HSSFCell myc=myr.createCell(0);
							myc.setCellValue(rs.getString(1));
							if(rs.getString(10)==null)//Partial Indicator= NULL
							{
								if(rs.getString(9)==null)//Manual Indicator= NULL
								{
									if(rs.getString(5)==null || rs.getString(5)=="Y")//Block Extract Indicator= ‘Y’ or NULL
									{
										if(rs.getString(6)!=null)//CFM start Date !=null
										{
											if(rs.getInt(12)<5)//Age Control shouldn’t be more than5
											{
												if(rs1.getString(3)!=null)//CFM start Date !=null in Query 2
												{
													if(rs2.getString(5)==null)
													{
														if(rs3.getString(2)==null && rs1.getString(9)=="UB-MOW")
														{
																																					
														HSSFCell myc1=myr.createCell(1);
														myc1.setCellValue("MA# is Mandatory for "+rs1.getString(9));
														}
														else
														{
															HSSFCell myc1=myr.createCell(1);
															myc1.setCellValue("Valid");
														}
													}
													else
													{
													HSSFCell myc1=myr.createCell(1);
													myc1.setCellValue("HIER_PNT_DESC_UTF8_TX value is "+ rs2.getString(5));
														
													}
												}
												else
												{
													HSSFCell myc1=myr.createCell(1);
													myc1.setCellValue("CFM Start date is present in Query 1 but not Present in Query 2 for "+ rs1.getString(9) +" country(US or UB-MOW Mandatory)");
													
												}
											}
											else
											{
												HSSFCell myc1=myr.createCell(1);
												myc1.setCellValue("Age control is "+rs.getString(12));
											}
										
										}
										else
										{
											HSSFCell myc1=myr.createCell(1);
											//myc1.setCellValue("CFM Start date is not present in Query 1");
											if(rs1.getString(3)==null)
											{
												myc1.setCellValue("CFM Start date is not present in Query 1 and Query 2 for "+ rs1.getString(9) +"country(US or UB-MOW Mandatory)");
											}
											
											else
											{
												myc1.setCellValue("CFM Start Date is not present in Query 1 but present in Query 2 for "+ rs1.getString(9) +"country(US or UB-MOW Mandatory)");
											}
											
										}
									}
									else
									{
										HSSFCell myc1=myr.createCell(1);
										myc1.setCellValue("Block Indicator is "+rs.getString(5));
									}
								
								}
								else
								{
									HSSFCell myc1=myr.createCell(1);
									myc1.setCellValue("Manual Indicator is "+rs.getString(9));
								}
									
							}
							else
							{
								HSSFCell myc1=myr.createCell(1);
								myc1.setCellValue("Partial Indicator is "+rs.getString(10));	
							}
							
							
						}
						FileOutputStream f1=new FileOutputStream(new File("C:\\ScannedData.xls"));
						wb.write(f1);
						
						f1.close();
						rs.close();
						rs1.close();
						rs2.close();
						rs3.close();
						
						}
						stmt.close();
						stmt1.close();
						stmt2.close();
						stmt3.close();
						
						con.close();
						
						
					}
					catch(Exception e1)
					{
						e1.printStackTrace();
					}
					
					break;*/
				case 3:
					/*y=LogInID.getText();
					z=password.getText();
					x=Sheet.getText();
					try {
						Connection con2 = DriverManager.getConnection("",y,z);
						Statement stmt2=con2.createStatement(); 
						ResultSet rs2=stmt2.executeQuery("");  
					} catch (SQLException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
				    }
				   break;*/
				
				
				}
			}
		});
		btnRun.setBounds(319, 516, 132, 33);
		contentPane.add(btnRun);
		
		password = new JPasswordField();
		password.setBounds(219, 166, 164, 33);
		contentPane.add(password);
	}
}
