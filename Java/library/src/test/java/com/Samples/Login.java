package com.Samples;

import java.io.IOException;

import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.library.generic.SAPGeneric;
import com.library.generic.SAPGuiClassName;

public class Login extends SAPGuiClassName {

	/*
	 * Declaring the ActiveXComponent Objects 
	 */
	ActiveXComponent SAPROTWr, GUIApp, Connection, Session, Obj,SAPGIRD;
	Dispatch ROTEntry;
	Variant Value, ScriptEngine;
	private Process p;	
	SAPGeneric sapGenric;
	/*
	 * Opens the SAP Session and returns the Active Session Object 
	 */
	@BeforeClass
	public void getSapSessionObject() throws InterruptedException{
		String cnt = "0";
		ComThread.InitSTA();
		
		//Opening the SAP Logon 
		try {
	        p = Runtime.getRuntime().exec("C:\\Program Files\\SAP\\FrontEnd\\SAPgui\\saplogon.exe");
	    } catch (IOException e) {
	        // TODO Auto-generated catch block
	        e.printStackTrace();
	    }
		Thread.sleep(5000);
		//-Set SapGuiAuto = GetObject("SAPGUI")---------------------------
		SAPROTWr = new ActiveXComponent("SapROTWr.SapROTWrapper");

		ROTEntry = SAPROTWr.invoke("GetROTEntry", "SAPGUI").toDispatch();
		//-Set application = SapGuiAuto.GetScriptingEngine------------
		ScriptEngine = Dispatch.call(ROTEntry, "GetScriptingEngine");
		GUIApp = new ActiveXComponent(ScriptEngine.toDispatch());

		SAPROTWr = new ActiveXComponent("SapROTWr.SapROTWrapper");
		//SAP Connection Name 
		Connection = new ActiveXComponent(
				GUIApp.invoke("OpenConnection","QR1 - Retail").toDispatch());
		
		//-Set session = connection.Children(0)-----------------------

		//Initialization for the SAPGeneric
		sapGenric  = new SAPGeneric(Connection);
		sapGenric.setSession(new ActiveXComponent(sapGenric.getSession().invoke("FindById", "wnd[0]").toDispatch()));
	}
	
	
	/*
	 * Logging to SAP GUI 
	 */
	@Test
	public void  LoginSAPGUI() throws InterruptedException
	{
		Thread.sleep(2000);
		sapGenric.SAPGUISetPasswordField("RSYST-BCODE", "username");
		sapGenric.SAPGUITextFieldSet("RSYST-BNAME", "password");
		sapGenric.SAPGUIEnter();
		
	}
	/*
	 * WorkBench exporting xml 
	 */
	public void exportStorePosTransactions(String storenumber, String dateOfTrans, String filelocation, String filename)
		    throws InterruptedException
		  {
		    Thread.sleep(5000L);
		    sapGenric.SAPGUIOKCodeFiled("okcd", "/n/posdw/xmlo");
		    sapGenric.SAPGUIEnter();
		    Thread.sleep(2000L);
		    sapGenric.SAPGUICTextFieldSendKeys("STORE", storenumber);
		    sapGenric.SAPGUICTextFieldSendKeys("DAY", dateOfTrans);
		    sapGenric.SAPGUIbtnClick("btn[8]");
		  }
     
}
