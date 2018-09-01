package com.library.generic;

import java.util.HashMap;
import java.util.Map;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class SAPGeneric {
	private ActiveXComponent Obj;
	private ActiveXComponent session;
	private ActiveXComponent parentSession;
	Variant[] arg = new Variant[2];
	static Map<String, String> SAPSearchIdMapObject = new HashMap<String, String>();
	
	public ActiveXComponent getSession() {
		return this.Obj;
	}

	public ActiveXComponent getParentSession() {
		return this.parentSession;
	}

	public void setSession(ActiveXComponent initSess) {
		this.Obj = initSess;
	}

	public void setParentSession(ActiveXComponent initSess) {
		this.parentSession = initSess;
	}

	public SAPGeneric(ActiveXComponent initSess) {
		this.Obj = new ActiveXComponent(initSess.invoke("Children", 0).toDispatch());
		this.parentSession = initSess;
	}

	public void SAPGUIEnter() {
		Obj.invoke("sendVKey", 0);
	}

	public boolean isObjectExist(String ObjName) {
		try {
			session = new ActiveXComponent(this.Obj.invoke("FindById", ObjName).toDispatch());
			return true;
		} catch (Exception ex) {
		}

		return false;
	}

	public boolean isWindowExist(String ObjName) {
		try {
			session = new ActiveXComponent(this.parentSession.invoke("FindById", ObjName).toDispatch());
			return true;
		} catch (Exception ex) {
		}

		return false;
	}

	public void SAPGUITextFieldSet(String lbl, String txtValue, String LabelClassName) {

		arg[0] = new Variant(lbl);
		arg[1] = new Variant(LabelClassName);
		//session = new ActiveXComponent(Obj.invoke("FindById", "wnd[0]/usr/pwdRSYST-BCODE").toDispatch());
	session = new ActiveXComponent(Obj.invoke("FindByName", arg).toDispatch());
		
		session.setProperty("text", txtValue);
	}
	public void SAPGUITextFieldSendKeys(String lbl, String KeyValue, String LabelClassName) {

		arg[0] = new Variant(lbl);
		arg[1] = new Variant(LabelClassName);
		session = new ActiveXComponent(Obj.invoke("FindByName", arg).toDispatch());
		session.invoke("setFocus");
		Obj.invoke("sendVKey",KeyValue);
	}


	/*
	 * SAPGUITextFieldGet Objective, to get the text of an text edit field
	 * Parameter Lbl :  name of the element, Parameter LableClassName: Gui Class name 
	 * returns String.
	 * created by Venkata Siva kumar 5/02/2018
	 */
	public String SAPGUITextFieldGet(String lbl, String LabelClassName) 
	{

		arg[0] = new Variant(lbl);
		arg[1] = new Variant(LabelClassName);
		session = new ActiveXComponent(Obj.invoke("FindByName", arg).toDispatch());
		return session.getProperty("text").toString();
	}

	/*
	 * SAPGUISelectOperation Objective, to perform select operation Parameter Lbl :
	 * name of the element Parameter LableClassName: Gui Class name created by
	 * Venkata Siva kumar 4/30/2018
	 */
	public void SAPGUISelectOperation(String lbl, String LabelClassName) {

		arg[0] = new Variant(lbl);
		arg[1] = new Variant(LabelClassName);
		session = new ActiveXComponent(Obj.invoke("FindByName", arg).toDispatch());
		Dispatch.call(session, "Select");
	}

	public ActiveXComponent SetObjectByname() {
		return new ActiveXComponent(Obj.invoke("FindByName", arg).toDispatch());

	}

	public void SAPGUITextFieldSet(String lbl, String txtValue) {
		
		SAPGUITextFieldSet(lbl, txtValue, SAPGuiClassName.GuiTextClassName);
	}
	
	public void SAPGUITextFieldSet(String lbl, String txtValue,int window) {
		ActiveXComponent wnd0 = getSession();
		setSession(new ActiveXComponent(getParentSession().invoke("FindById", "wnd["+window+"]").toDispatch()));
		SAPGUITextFieldSet(lbl, txtValue, SAPGuiClassName.GuiTextClassName);
		setSession(wnd0);
	}
/*
 	 * SAPGUICTextFieldSendKeys Objective, Sendkeys in text field :
	 * name of the element Parameter Label: KeyValue(ex :F1 -> 01 02->F2  
	 * 03 -> F3 04 ->F4 05 -> F5 06 -> F6 07 ->F7 08 ->F8 ) name created by
	 * Venkata  06/01/2018
	 *  F1 -> 01 02->F2  03 -> F3 04 ->F4 05 -> F5 06 -> F6 07 ->F7 08 ->F8 

 */
	public void SAPGUICTextFieldSendKeys(String lbl, String KeyValue) {
		SAPGUITextFieldSendKeys(lbl, KeyValue, SAPGuiClassName.GuiCTextClassName);
	}
	/*
 	 * SAPGUITextFieldSendKeys Objective, Sendkeys in text field :
	 * name of the element Parameter Label: KeyValue(ex :F1 -> 01 02->F2  
	 * 03 -> F3 04 ->F4 05 -> F5 06 -> F6 07 ->F7 08 ->F8 ) name created by
	 * Venkata  06/01/2018
	 *  F1 -> 01 02->F2  03 -> F3 04 ->F4 05 -> F5 06 -> F6 07 ->F7 08 ->F8 

 */
	public void SAPGUITextFieldSendKeys(String lbl, String KeyValue) {
		SAPGUITextFieldSendKeys(lbl, KeyValue, SAPGuiClassName.GuiTextClassName);
	}
	public void SAPGUISetCTextField(String lbl, String txtValue) {
		SAPGUITextFieldSet(lbl, txtValue, SAPGuiClassName.GuiCTextClassName);
	}

	public void SAPGUISetCTextField(String lbl, String txtValue, int window) {
		ActiveXComponent wnd0 = getSession();
		setSession(new ActiveXComponent(getParentSession().invoke("FindById", "wnd["+window+"]").toDispatch()));
		SAPGUITextFieldSet(lbl, txtValue, SAPGuiClassName.GuiCTextClassName);
		setSession(wnd0);
	}
	
	public void SAPGUISetPasswordField(String lbl, String txtValue) {
		SAPGUITextFieldSet(lbl, txtValue, SAPGuiClassName.GuiPasswordClassName);
	}

	public void SAPGUIClick(String lbl, String LabelClassName) throws InterruptedException {
		arg[0] = new Variant(lbl);
		arg[1] = new Variant(LabelClassName);
		session = new ActiveXComponent(this.Obj.invoke("FindByName", arg).toDispatch());
		Dispatch.call(session, "press");
		Thread.sleep(1000);
	}

	public void SAPGUIRadiobtnClick(String lbl) {
		SAPGUISelectOperation(lbl, SAPGuiClassName.GuiRadioButtonClassName);
	}

	public void SAPGUIbtnClick(String lblname) throws InterruptedException {
		SAPGUIClick(lblname, SAPGuiClassName.GuiButtonClassName);
	}
	
	public void SAPGUIbtnClick(String lblname, int window) throws InterruptedException {
		ActiveXComponent wnd0 = getSession();
		setSession(new ActiveXComponent(getParentSession().invoke("FindById", "wnd["+window+"]").toDispatch()));
		SAPGUIClick(lblname, SAPGuiClassName.GuiButtonClassName);
		setSession(wnd0);
	}

	public void SAPGUIOKCodeFiled(String name, String value) {
		SAPGUITextFieldSet(name, value, SAPGuiClassName.GuiOkCodeFieldClassName);
	}

	/*
	 * Function getSAPObjectIDHelperMethod is helper method to call the
	 * getSAPObjectIDByComparingTheProperty_Recursive returns the Id of the Object
	 * if found else returns null
	 */
	public String getSAPObjectIDHelperMethod(ActiveXComponent Session, String PrimaryPropertyName,
			String PrimaryPropertyValue, String SecondaryPropertyName, String SecondaryPropertyValue)

	{
		SAPSearchIdMapObject.clear();
		ActiveXComponent st = getSAPObjectIDByComparingTheProperty_Recursive(Session, PrimaryPropertyName,
				PrimaryPropertyValue, SecondaryPropertyName, SecondaryPropertyValue);
		if (!SAPSearchIdMapObject.containsValue(""))
			return SAPSearchIdMapObject.get("ID");
		else
			return "";

	}

	/*
	 * Generic method :getSAPObjectIDByComparingTheProperty_Recursive Objective : To
	 * search the Sap object based on the properties sent as parameter. Parameters
	 * Session : SAP Object, PrimaryPropertyName : Name of the property ,
	 * PrimaryPropertyValue : Value of the property First preference is given to the
	 * primary parameters if primary property matched only then the secondary
	 * properties are checked.. Secondary parameter name and value can be blank.
	 * Function in response it will set the Hash map member of the class with the
	 * ID, if object found We can use this method to search the partial ID as well.
	 * it is recursive method
	 */

	@SuppressWarnings({ "deprecation", "static-access", "null" })
	public ActiveXComponent getSAPObjectIDByComparingTheProperty_Recursive(ActiveXComponent Session,
			String PrimaryPropertyName, String PrimaryPropertyValue, String SecondaryPropertyName,
			String SecondaryPropertyValue) {

		int count = 0;
		boolean foundXpath = false;

		String xpath = "";
		// boolean recFlag = true;
		ActiveXComponent childObject = null, currentObj = null;
		Dispatch disp = Session.invoke("Children").toDispatch();
		count = disp.call(disp, "count").toInt();
		for (int i = 0; i < count; i++) {
			currentObj = new ActiveXComponent(Session.invoke("Children", i).toDispatch());
			String tempXpath = currentObj.getProperty(PrimaryPropertyName).toString();
			String SubTypeNameActual = tempXpath;

			if (tempXpath.toString().length() > 1) {
				if (tempXpath.toString().contains(PrimaryPropertyValue) && SecondaryPropertyName.length() > 1)
					SubTypeNameActual = currentObj.getProperty(SecondaryPropertyName).toString();
			}
			if (tempXpath.toString().contains(PrimaryPropertyValue)
					&& SubTypeNameActual.toString().contains(SecondaryPropertyValue)) {
				xpath = currentObj.getProperty("ID").toString();
				xpath = xpath.substring("/app/con[0]/ses[0]/".length());
				System.out.println("Xpath found " + xpath);
				// int sizeid = SAPSearchIdMapObject.keySet().size();
				SAPSearchIdMapObject.put("ID", xpath.trim());
				return currentObj;

			}
			if (currentObj.getProperty("ContainerType").toBoolean() && !foundXpath) {
				childObject = getSAPObjectIDByComparingTheProperty_Recursive(currentObj, PrimaryPropertyName,
						PrimaryPropertyValue, SecondaryPropertyName, SecondaryPropertyValue);

			}
		}
		if (childObject != null)
			currentObj = childObject;
		return currentObj;

	}

	/*
	 * SAPGUISelectTab Objective, Select the Tab by name Parameter lbl : Tab Name
	 * created by Venkata Siva kumar 4/30/2018
	 */
	public void SAPGUISelectTab(String lbl) 
	{
		SAPGUISelectOperation(lbl, SAPGuiClassName.GuiTabClassName);
	}

	/*
	 * GuiComboBoxSelectItem Objective, to select the combo item by value 
	 * created by Venkata Siva kumar 4/30/2018
	 * 
	 */
	public void GuiComboBoxSelectItem(String lbl, String ItemToSelect) {
		arg[0] = new Variant(lbl);
		arg[1] = new Variant(SAPGuiClassName.GuiComboBoxClassName);
		session = new ActiveXComponent(Obj.invoke("FindByName", arg).toDispatch());
		session.setProperty("value", ItemToSelect);
	}

	/*
	 * SAPGUICheckBoxOnOrOff Objective, to select or unselect the check box
	 * parameter lbl : Name of the check box parameter OnOrOff : Pass true to select
	 * the radio button and false to deselect the radio button 
	 * created by Venkata Siva kumar
	 * 4/30/2018
	 * 
	 */
	public void SAPGUICheckBoxOnOrOff(String lbl, Boolean OnOrOff) {
		arg[0] = new Variant(lbl);
		arg[1] = new Variant(SAPGuiClassName.GuiCheckBoxClassName);
		session = new ActiveXComponent(this.Obj.invoke("FindByName", arg).toDispatch());
		if (OnOrOff)
			session.setProperty("selected", -1);
		else
			session.setProperty("selected", 0);

	}
	
	 

	/* SAPGUITableControlGetVisibleRowCount 
	 * Objective - Function fetches the visible row count of the Table control object 
	 * Parameter  lbl - Name of the TableControl, ColumnNmae - Name of the column 
	 * returns -1 if the column is not found else it returns the column number
	 * Created by Venkata Siva kumar 
	 */

	public int SAPGUITableControlGetVisibleRowCount(String lbl) 
	{
		int count  = 0;
		 
			arg[0] = new Variant(lbl);
			arg[1] = new Variant(SAPGuiClassName.GuiTableControlClassName);			 
			session = new ActiveXComponent(getSession().invoke("FindByName", arg).toDispatch());			
			count = Dispatch.call(session, "VisibleRowCount").toInt();
		return count;
		
	}

	/* SAPGUITableControlGetColumnIndex 
	 * Objective - Function fetches the column number for an Table control object 
	 * Parameter  lbl - Name of the TableControl, ColumnNmae - Name of the column 
	 * returns -1 if the column is not found else it returns the column number
	 * Created by Venkata Siva kumar 
	 */
	@SuppressWarnings("deprecation")
	public int SAPGUITableControlGetColumnIndex(String lbl, String ColumName ) 
	{
			String str = null;
			ActiveXComponent currentObj;
			arg[0] = new Variant(lbl);
			arg[1] = new Variant(SAPGuiClassName.GuiTableControlClassName);				 
			session = new ActiveXComponent(getSession().invoke("FindByName",  arg).toDispatch());	
			Dispatch dis = Dispatch.call(session, "Columns").toDispatch();
			int count = Dispatch.call(dis, "count").toInt(); 			 
			Dispatch[] arr = new Dispatch[count];
			for(int i=0;i<count;i++)
			{
				currentObj = new ActiveXComponent(session.invoke("Columns", i).toDispatch());
				if( currentObj.getProperty("title").toString().trim().equals(ColumName))
				{
					return i;
				}
			}			
			
		return -1;
		
	}
	
	/* SAPGUITableControlGetCellValue 
	 * Objective - Function fetches the value of the cell value 
	 * Parameter  lbl - Name of the TableControl, ColumnNmae - Name of the column, rowid - row number
	 * returns -1 if the column is not found else it returns the cell value
	 * Created by Venkata Siva kumar 
	 */	
	@SuppressWarnings("deprecation")
	public String SAPGUITableControlGetCellValue(String lbl, int rowid, String ColumnName ) 
	{
			String str = null;
			ActiveXComponent currentObj =SAPGUITableControlGetCellObject(lbl,rowid,ColumnName);		
			str= currentObj.getProperty("text").toString();					 
			return str;
		
	}
	
	/* SAPGUITableControlGetCellObject 
	 * Objective - Function fetches the value of the cell value 
	 * Parameter  lbl - Name of the TableControl, ColumnNmae - Name of the column, rowid - row number
	 * returns null if the column is not found else it returns the cell object 
	 * Created by Venkata Siva kumar 
	 */	
	@SuppressWarnings("deprecation")
	public ActiveXComponent SAPGUITableControlGetCellObject(String lbl, int rowid, String ColumnName ) 
	{
			String str = null;
			ActiveXComponent currentObj=null;
			arg[0] = new Variant(lbl);
			arg[1] = new Variant(SAPGuiClassName.GuiTableControlClassName);				 
			session = new ActiveXComponent(getSession().invoke("FindByName",  arg).toDispatch());
			int Columnindex = SAPGUITableControlGetColumnIndex(lbl, ColumnName);
			currentObj = new ActiveXComponent(session.invoke("GetCell", rowid, Columnindex ).toDispatch());			
			return currentObj;
		
	}
	 	
	/* SAPGUITableControlSubOperations 
	 * Objective - Function performs the subtype operations on the table cell based on the cell type by setting the focus on cell
	 * Parameter  lbl - Name of the TableControl, ColumnName - Name of the column, rowid - row number 
	 * Parameter SubTaskClassName - Send the class name of the control (refer SAPGuiClassName) 
	 * Value - depends upon the type of the control refer the below conditions 
	 * 		for "GuiTextField" and "GuiCTextField" send the string value to be set
	 * 		for "GuiCheckBox" need to send boolan value true/false in the form of string 
	 * 		for "GuiComboBox" item value to be selected 
	 * 		for "GuiButton" clicks on button and value field is not considered and it can be blank or null 
	 * returns true if we are able to perform the action successfully else returns false 
	 * Created by Venkata Siva kumar 
	 */	
	@SuppressWarnings("deprecation")
	public boolean SAPGUITableControlSubOperations(String lbl, int rowid, String ColumnName, String SubTaskClassName, String Value ) 
	{
		String str = null;
		ActiveXComponent currentObj =SAPGUITableControlGetCellObject(lbl,rowid,ColumnName);	
		if(!(currentObj.getProperty("Type").toString().trim().equals(SubTaskClassName)))
		{			
			System.out.println( "Table  "+ lbl + " cell class name is "+currentObj.getProperty("Type").toString().trim() +" and its not matching with the expected"+SubTaskClassName);
			return false;
		}	
		currentObj.invoke("setFocus");
		 switch (SubTaskClassName) 
		 {
         	case "GuiTextField" :
         		currentObj.setProperty("text",Value  );	
         		break;	
         			
         	case "GuiCTextField" :
         		currentObj.setProperty("text",Value  );	
    			break;
    			
         	case "GuiCheckBox" :         		
         		if (Boolean.parseBoolean(Value))
         			currentObj.setProperty("selected", -1);
    			else
    				currentObj.setProperty("selected", 0);         		
         		break;
         	
         	case "GuiComboBox" :    
         		currentObj.setProperty("value", Value);	
         		break;
         		
         	case "GuiButton" :
         		Dispatch.call(currentObj, "press");
         		break;         	
         		
         	default :
         		System.out.println("Sub task " +SubTaskClassName +"is not handled ");
         		return false;
		 }
		return true;
		
		
	
	}
	
	/* SAPGUITableControlSetCellObject 
	 * Objective - Function fetches the value of the cell value 
	 * Parameter  lbl - Name of the TableControl, ColumnName - Name of the column, rowid - row number
	 * ValueToBeSet - new value we are trying to set  
	 * Created by Venkata Siva kumar 
	 */	
	@SuppressWarnings("deprecation")
	public void SAPGUITableControlSetCellValue(String lbl, int rowid, String ColumnName, String ValueToBeSet ) 
	{
		String str = null;
		ActiveXComponent currentObj =SAPGUITableControlGetCellObject(lbl,rowid,ColumnName);		
		currentObj.setProperty("text",ValueToBeSet  );	
		
	}

	/* SAPGUITableControlSetFocusOnCell 
	 * Objective - Function sets focus on the cell 
	 * Parameter  lbl - Name of the TableControl, ColumnNmae - Name of the column, rowid - row number	
	 * Created by Venkata Siva kumar 
	 */	
	@SuppressWarnings("deprecation")
	public void SAPGUITableControlSetFocusOnCell(String lbl, int rowid, String ColumnName ) 
	{
			String str = null;
			ActiveXComponent currentObj =SAPGUITableControlGetCellObject(lbl,rowid,ColumnName);		
			currentObj.invoke("setFocus");
		
	}


/*SAPGUITreeExpandNode
   * Objective - To expand node in a tree
   * parameter  : id, treename and nodevalue
   * created by Venkata Siva Kumar    
   */
  
	public void SAPGUITreeExpandNode(String id, String treeName, String nodeValue) throws Exception {
	Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
    Dispatch dis = Dispatch.call(this.Obj, "GetAllNodeKeys").toDispatch();
    Variant[] var = new Variant[2];
    int count = Dispatch.call(dis, "count").toInt();
    for (int i = 0; i < count; i++)
    {
      var[0] = Dispatch.call(this.Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(i) });
      var[1] = new Variant(treeName);
      if (Dispatch.call(Obj, "GetNodeTextByKey", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(i) }) }).toString().contains(nodeValue))
      {
        Obj.invoke("expandNode", new Variant[] { var[0] });
        break;
      }
    }
    
    Thread.sleep(1000L);
  }
  
	/*SAPGUITreeExpandAndDoubleClickNode
	   * Objective - To expand and double click node in a tree
	   * parameter  : id, treename and nodevalue
	   * created by Venkata Siva Kumar    
	   */
  public void SAPGUITreeExpandAndDoubleClickNode(String id,String treeName, String nodeValue) throws Exception
  {
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
    Dispatch dis = Dispatch.call(Obj, "GetAllNodeKeys").toDispatch();
    Variant[] var = new Variant[2];
    int count = Dispatch.call(dis, "count").toInt();
    for (int i = 0; i < count; i++)
    {
      var[0] = Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(i) });
      var[1] = new Variant(treeName);
      if (Dispatch.call(Obj, "GetNodeTextByKey", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(i) }) }).toString().contains(nodeValue))
      {
        Obj.invoke("expandNode", new Variant[] { var[0] });
        Thread.sleep(2000L);
        Obj.invoke("doubleClickItem", var);
        break;
      }
    }
    
    Thread.sleep(2000L);
  }
  
  /*SAPGUITreeSelectNode
   * Objective - To select/activate node in a tree
   * parameter  : id, treename and nodevalue
   * created by Venkata Siva Kumar    
   */
  public void SAPGUITreeSelectNode(String id, String treeName, String nodeValue) throws Exception
  {
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
    Dispatch dis = Dispatch.call(Obj, "GetAllNodeKeys").toDispatch();
    Variant[] var = new Variant[2];
    int count = Dispatch.call(dis, "count").toInt();
    for (int i = 0; i < count; i++)
    {
      var[0] = Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(i) });
      var[1] = new Variant(treeName);
      if (Dispatch.call(Obj, "GetNodeTextByKey", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(i) }) }).toString().contains(nodeValue))
      {
    	  Dispatch.call(Obj, "selectItem", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(i) }), treeName });
    	  //Obj.invoke("clickLink", var);
        break;
      }
    }
    
    Thread.sleep(1000L);
  }
  
  /*SAPGUITreeSetRowAndDoubleClickCell
   * Objective - To activate node in a tree
   * parameter  : id, treename and cellValue
   * created by Venkata Siva Kumar    
   */
  public void SAPGUITreeSetRowAndDoubleClickCell(String id, String columnName, String cellValue) throws Exception
  {
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
    int count = ActiveXComponent.call(Obj, "RowCount").toInt();
    for (int i = 0; i < count; i++)
    {
      Obj.setProperty("firstVisibleColumn", columnName);
      Obj.setProperty("firstVisibleRow", String.valueOf(i));
      if (Dispatch.call(Obj, "GetCellValue", new Object[] { Integer.valueOf(i), columnName }).toString().contains(cellValue))
      {
        Obj.setProperty("currentCellColumn", columnName);
        Obj.setProperty("currentCellRow", i);
        Obj.invoke("doubleClickCurrentCell");
        break;
      }
    }
    Thread.sleep(1000L);
  }
  
  /*SAPGUITreeClickNode
   * Objective - To select/activate link in a tree
   * parameter  : id, treename and nodevalue,gridValCheck
   * created by Venkata Siva Kumar    
   */
  public void SAPGUITreeClickNode(String id, String tree, String val, String gridValCheck) throws Exception
  {
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
    Dispatch dis = Dispatch.call(Obj, "GetAllNodeKeys").toDispatch();
    Variant[] var = new Variant[2];
    int count = Dispatch.call(dis, "count").toInt();
    
    for (int k = 0; k < count; k++)
    {
      var[0] = Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) });
      var[1] = new Variant(tree);
      
      String text = Dispatch.call(Obj, "GetItemText", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) }), tree }).toString();
      if (text.contains(val))
      {
        Dispatch.call(Obj, "selectItem", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) }), tree });
        Thread.sleep(1000L);
        Dispatch.call(Obj, "ensureVisibleHorizontalItem", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) }), tree });
        Dispatch.call(Obj, "ClickLink", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) }), tree });
        Thread.sleep(1000L);
        break;
      }
    }
  }
  
  /*SAPGUITreeLinkExpandNode
   * Objective - To expand link in a tree
   * parameter  : id, treename and nodevalue
   * created by Venkata Siva Kumar    
   */
  public void SAPGUITreeLinkExpandNode(String id, String tree, String val) throws Exception{
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
    Dispatch dis = Dispatch.call(Obj, "GetAllNodeKeys").toDispatch();
    Variant[] var = new Variant[2];
    int count = Dispatch.call(dis, "count").toInt();
    
    for (int k = 0; k < count; k++)
    {
      var[0] = Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) });
      var[1] = new Variant(tree);
      
      String text = Dispatch.call(Obj, "GetItemText", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) }), tree }).toString();
      if (text.contains(val))
      {
        Dispatch.call(Obj, "selectItem", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) }), tree });
        Dispatch.call(Obj, "expandNode", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { Integer.valueOf(k) }), tree });
        break;
      }
    }
  }
  
  /*SAPGUITreeGetNodeText
   * Objective - To get node text in a tree
   * parameter  : id, treename and key
   * created by Venkata Siva Kumar    
   */
  
  public String SAPGUITreeGetNodeText(String id, String tree, String key) throws Exception
  {
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
    Dispatch dis = Dispatch.call(Obj, "GetAllNodeKeys").toDispatch();
    Variant[] var = new Variant[2];
    int count = Dispatch.call(dis, "count").toInt();
    String text = "";
    var[0] = Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { key });
    var[1] = new Variant(tree);
    text = Dispatch.call(Obj, "GetItemText", new Object[] { Dispatch.call(Obj, "GetAllNodeKeys", new Object[] { key }), tree }).toString();
    return text;
  }
  
  /*SAPGUIGridGetCellText
   * Objective - to get  the cell value from the grid 
   * parameter  : idstr - PArtial id of the Grid Control, String ColumnName - Name of the column , rowId - row index
   * returns the cell value if the row and column is not found exception is thrown
   * created by Venkata Siva kumar    
   */
  
  public String SAPGUIGridGetCellText(String idstr, String columnName, int rowid ) throws Exception
  {
	  
      String val = "";
      String id = getSAPObjectIDHelperMethod(getSession(), "ID", idstr, "", "");
	  session= new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());		    
      val = Dispatch.call(session, "GetCellValue", new Object[] { Integer.valueOf(rowid), columnName }).toString();       
      return val;
    
  }
  
  public double SAPGUIGridGetCellInt(String column) throws Exception
  {
   
      int numberOfItems = 0;
      double retVal = 0.0D;
      String val = "";
      numberOfItems = Integer.parseInt(Dispatch.call(Obj, "RowCount").toString());
      for (int j = 0; j < numberOfItems - 1; j++)
      {
        val = Dispatch.call(Obj, "GetCellValue", new Object[] { Integer.valueOf(j), column }).toString();
        if (!val.equals(""))
          retVal += Double.parseDouble(val.replace(",", ".").replace("-", ""));
        else
        	retVal = 0.0D;
      }
      return retVal;
  }
  
  /*SAPGUIGridSelectCell
   * Objective - To select cell in a grid
   * parameter  : partial id, row and column
   * created by Venkata Siva Kumar    
   */
  public void SAPGUIGridSelectCell(String idstr, int row, String column) throws Exception {
	 
		  String id = getSAPObjectIDHelperMethod(this.Obj, "ID", idstr, "", "");
		  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
		  Dispatch.call(Obj, "setCurrentCell", new Object[] { row, column });
  }
  
  /*SAPGuiToolbarSelectMenuItem
   * Objective - To select item from menu
   * parameter  : itemToSelect
   * created by Venkata Siva Kumar    
   */
  public void SAPGuiToolbarSelectMenuItem(String itemToSelect) throws Exception{
	 
		  String id = getSAPObjectIDHelperMethod(this.Obj, "Name", itemToSelect, "type", "GuiMenu");
		  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
		  Dispatch.call(session, "select");
  }
  
  /*SAPGuiEditTxtAreaCheck
   * Objective - To validate the partial text and return boolean
   * parameter  : name, textToValidate
   * created by Venkata Siva Kumar    
   */
  public boolean SAPGuiEditTxtAreaCheck(String name,String textToValidate) throws Exception{
	  boolean flag = false;
	 
		  String id = getSAPObjectIDHelperMethod(this.Obj, "Name", name, "type", "GuiCTextField");
		  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
		  String textUI = session.getProperty("text").toString();
		  if(textUI.contains(textToValidate))
			 flag =true;
		  else
			  flag = false;
	return flag;
  }
  
  /*SAPGuiLabelDoubleClick
   * Objective - To select label and double click
   * parameter  : name, label
   * created by Venkata Siva Kumar    
   */
  public boolean SAPGuiLabelDoubleClick(String objName,String label) throws Exception{
	  boolean flag = false;
	  
		  arg[0] = new Variant(objName);
		  arg[1] = new Variant(label);
		  session = new ActiveXComponent(getSession().invoke("FindByName", arg).toDispatch());
		  String textUI = session.getProperty("text").toString();
		  if(textUI.contains(label))
		  { 
			  Dispatch.call(session, "setFocus");
			  session.setProperty("CaretPosition", 0);
			  getSession().invoke("sendVKey",2);
		  }
	return flag;
  }
  
  /*SAPGuiGridCellNonZeroCheck
   * Objective - To select label and double click
   * parameter  : name, label
   * created by Venkata Siva Kumar    
   */
  public boolean SAPGuiGridCellNonZeroCheck(String idstr,String row, String column) throws Exception{
	    boolean flag = false;
	    String id = getSAPObjectIDHelperMethod(this.Obj, "ID", idstr, "", "");
		  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
	        String val = Dispatch.call(Obj, "GetCellValue", new Object[] { Integer.valueOf(row), column }).toString();
	        if (Double.parseDouble(val) !=0)
	          flag = true;
	        else
	        	flag = false;
	        return flag;
	  }
  
  /*SAPGuiGridCellDoubleClick
   * Objective - To double click on cell in a grid
   * parameter  : partial id, row and column
   * created by Venkata Siva Kumar    
   */
  public void SAPGuiGridCellDoubleClick(String idstr, int row, String column) throws Exception {
	 
		  String id = getSAPObjectIDHelperMethod(this.Obj, "ID", idstr, "", "");
		  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
		  Dispatch.call(Obj, "setCurrentCell", new Object[] { row, column });
		  Dispatch.call(Obj, "DoubleClickCurrentCell");
  }
  
  /*SAPGuiGridCellValueCheck
   * Objective - To check the value in a grid cell
   * parameter  : partial id, row and column, valueToCheck
   * created by Venkata Siva Kumar    
   */
  public boolean SAPGuiGridCellValueCheck(String idstr,String row,String column,String valueToCheck) throws Exception
  {
	  boolean flag = false;
	  String id = getSAPObjectIDHelperMethod(this.Obj, "ID", idstr, "", "");
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
       String val = Dispatch.call(Obj, "GetCellValue", new Object[] { Integer.valueOf(row), column }).toString();
        if (val.equals(valueToCheck))
          flag = true;
        else
        	flag = false;
      
      return flag;
  }
  
  /*SAPGuiLabelCheck
   * Objective - To validate the label content
   * parameter  : name, label
   * created by Venkata Siva Kumar    
   */
  public boolean SAPGuiLabelCheck(String objName,String label) throws Exception{
	  boolean flag = false;
	  
		  arg[0] = new Variant(objName);
		  arg[1] = new Variant(label);
		  session = new ActiveXComponent(getSession().invoke("FindByName", arg).toDispatch());
		  String textUI = session.getProperty("text").toString();
		  if(textUI.contains(label))
		  { 
			 flag = true;
		  }
	return flag;
  }
  
  /*SAPGuiLabelSetFocus
   * Objective - To set focus on label
   * parameter  : name, label
   * created by Venkata Siva Kumar    
   */
  public boolean SAPGuiLabelSetFocus(String objName,String label) throws Exception{
	  boolean flag = false;
	  
		  arg[0] = new Variant(objName);
		  arg[1] = new Variant(label);
		  session = new ActiveXComponent(getSession().invoke("FindByName", arg).toDispatch());
		  String textUI = session.getProperty("text").toString();
		  if(textUI.contains(label))
		  { 
			  Dispatch.call(session, "setFocus");
			  session.setProperty("CaretPosition", 0);
		  }
	return flag;
  } 
  
  /*SAPGuiGridCellNonEmptyCheck
   * Objective - To check if grid cell is not empty
   * parameter  : partial id, row and column, valueToCheck
   * created by Venkata Siva Kumar    
   */
  public boolean SAPGuiGridCellNonEmptyCheck(String idstr,String row,String column,String valueToCheck) throws Exception
  {
	  boolean flag = false;
	  String id = getSAPObjectIDHelperMethod(this.Obj, "ID", idstr, "", "");
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
       String val = Dispatch.call(Obj, "GetCellValue", new Object[] { Integer.valueOf(row), column }).toString();
        if (!val.equals(""))
          flag = true;
        else
        	flag = false;
      
      return flag;
  }
  
  /*SAPGuiButtonClickTillLabelValueChange
   * Objective - To click on button and validate the label content till value changes as expected
   * parameter  : btnName, labelID, labelToCheck(text)
   * created by Venkata Siva Kumar    
   */
  public boolean SAPGuiButtonClickTillLabelValueChange(String btnName,String labelID,String labelToCheck) throws Exception
  {
	  String textUI = "";
	  boolean flag = false;
	  for(int i=0;i<500;i++)
	  {
		  SAPGUIbtnClick(btnName);
		  session = this.Obj;
		  String id = getSAPObjectIDHelperMethod(this.Obj, "ID", labelID, "Type", "GuiLabel");
		  Obj = new ActiveXComponent(getSession().invoke("FindById", id).toDispatch());
		  textUI = Obj.getProperty("text").toString();
		  if(textUI.contains(labelToCheck))
		  { 
			  flag = true;
			  break;
		  }
		  Thread.sleep(3000);
		  this.Obj = session;
	  }
	  if(flag)
	  {
		  this.Obj = session;
		  return true;
	  }
	  else
	  {
		  this.Obj = session;
		  return false;
	  }
  }
  
  /*SAPGuiGridGetCellValueplus
   * Objective - To fetch value from cell in grid and add 1 to it
   * parameter  : partial id, row and column
   * created by Venkata Siva Kumar    
   */
  public Integer SAPGuiGridGetCellValueplus(String idstr,String row,String column) throws Exception
  {
	  Integer finalVal = 0;
	  String id = getSAPObjectIDHelperMethod(this.Obj, "ID", idstr, "", "");
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
       String val = Dispatch.call(Obj, "GetCellValue", new Object[] { Integer.valueOf(row), column }).toString();
        if (!val.equals(""))
          finalVal = Integer.parseInt(val)+1;
        
      return finalVal;
  }

  /*SAPGuiGridGetRowCount
   * Objective - To get row count for a grid
   * parameter  : partial id
   * created by Venkata Siva Kumar    
   */
  public int SAPGuiGridGetRowCount(String partialID) throws Exception
  {	    
	  String id = getSAPObjectIDHelperMethod(this.Obj, "ID", partialID, "", "");
	  Obj = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
      int rowCount = Dispatch.call(Obj, "RowCount").toInt();
      return rowCount;    
  }
  
  
  /*SAPGUIElementFound
	   * Objective - To check weather the object found
	   * elementPropertyName  : Element property Name (such as name,type etc)
	   * elementPropertyValue : Value of the property 
	   * addPropertyName  : additional  Element property Name
	   * addPropertyValue : additional element property value 
	   * primary parameters if primary property matched only then the additional properties are matched 
	   * properties are checked.. additional parameter name and value can be blank.
	   * returns true if the object found else false
	   * created by Venkata Siva kumar    
   */
	public boolean SAPGUIElementFound(String elementPropertyName , String elementPropertyValue, String addPropertyName, String addPropertyValue)
	{
		 
			String id = getSAPObjectIDHelperMethod(this.Obj, elementPropertyName , elementPropertyValue, addPropertyName,addPropertyValue);
			if(id.length()> 1 )
					return true;			
			else
				return false;		 
	
	}
	
	/*SAPGUISelectOrDsselectRowInTableControl
	   * Objective -to select the row in the table control
	   * parameter  : Element Name , row number to select
	   * parameter selectordeselect send true to select the row and false to deselect the row
	   * created by Venkata Siva kumar    
	  */
	public void SAPGUISelectOrDsselectRowInTableControl(String Lbl, int Rowid , boolean selectordeselect) 
	{ 
		ActiveXComponent currentObj=null;
		arg[0] = new Variant(Lbl);
		arg[1] = new Variant(SAPGuiClassName.GuiTableControlClassName);				 
		session = new ActiveXComponent(getSession().invoke("FindByName",  arg).toDispatch());		
		currentObj = new ActiveXComponent(session.invoke("getAbsoluteRow", Rowid ).toDispatch());			
		currentObj.setProperty("selected", selectordeselect);			 
	}
	
	/*SAPGUISelectOrDsselectMultipleRowInTableControl
	   * Objective -to select the row in the table control
	   * parameter  : Element Name , Rowlist - List of rows to select separated by Pipe
	   * parameter selectordeselect pass true to select the row and false to deselect the row
	   * created by Venkata Siva kumar    
	  */
	public void SAPGUISelectOrDsselectMultipleRowInTableControl(String Lbl, String RowList , boolean selectordeselect) 
	{ 
		ActiveXComponent currentObj=null;
		arg[0] = new Variant(Lbl);
		arg[1] = new Variant(SAPGuiClassName.GuiTableControlClassName);				 
		session = new ActiveXComponent(getSession().invoke("FindByName",  arg).toDispatch());	
		for (String str : RowList.split("\\|"))
		{
			 
			currentObj = new ActiveXComponent(session.invoke("getAbsoluteRow", Integer.parseInt(str) ).toDispatch());			
			currentObj.setProperty("selected", selectordeselect);			 
		}		
		
	}
	
	
	
	 /*SAPGuiLabelcomp
	   * Objective - To compare the label text on UI with expected
	   * parameter  : name, label
	   * created by Venkata Siva Kumar    
	   */
	  public boolean SAPGuiLabelcomp(String objName,String label) throws Exception{
		  boolean flag = false;
		  	try{
			  arg[0] = new Variant(objName);
			  arg[1] = new Variant(label);
			  session = new ActiveXComponent(getSession().invoke("FindByName", arg).toDispatch());
			  String textUI = session.getProperty("text").toString();
			  if(textUI.equals(label))
			  { 
				  return true;
			  }
		  	}
		  	catch (Exception e)
		  	{
		  		return false;
		  	}
		  
		return flag;
	  } 
	  
	  /*SAPGuiGridCheckValDoubleClickCell
	   * Objective - To double click on cell in a grid
	   * parameter  : partial id, column and valToCheck
	   * created by Venkata Siva Kumar    
	   */
	  public boolean SAPGuiGridCheckValDoubleClickCell(String idstr, String column,String valueToCheck) throws Exception {
		  boolean flag = false;
		  String id = getSAPObjectIDHelperMethod(getSession(), "ID", idstr, "", "");
		  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
		  int rowCount = Dispatch.call(session, "RowCount").toInt();
		  for(int i=0;i<rowCount;i++)
		  {
			  String val = Dispatch.call(session, "GetCellValue", new Object[] { Integer.valueOf(i), column }).toString();
			  if (val.equals(valueToCheck))
			  {
				  Dispatch.call(session, "setCurrentCell", new Object[] { i, column });
				  Dispatch.call(session, "DoubleClickCurrentCell");
				  flag = true;
				  break;
			  }
			  else
				  flag = false;
		  }
		  return flag;
	  }
	  
	  /*SAPGUITableControlGetColumnNames
	   * Objective - to fetch all the column names of the table control 
	   * parameter  : Control Name 
	   * returns : Column names separated by '|' and returns null if no columns found 
	   * created by Venkata Siva kumar     
	   */
	  public String SAPGUITableControlGetColumnNames(String lbl)
	  {
		  	String str = null;
			ActiveXComponent currentObj;
			arg[0] = new Variant(lbl);
			arg[1] = new Variant(SAPGuiClassName.GuiTableControlClassName);				 
			session = new ActiveXComponent(getSession().invoke("FindByName",  arg).toDispatch());	
			Dispatch dis = Dispatch.call(session, "Columns").toDispatch();
			int count = Dispatch.call(dis, "count").toInt(); 			 
			Dispatch[] arr = new Dispatch[count];
			String ColumnNames =null;
			for(int i=0;i<count;i++)
			{
				currentObj = new ActiveXComponent(session.invoke("Columns", i).toDispatch());
				if(i==0)
					ColumnNames =  currentObj.getProperty("title").toString().trim() ;	
				else
					ColumnNames = ColumnNames +  "|" + currentObj.getProperty("title").toString().trim() ;	
				
			}			
			
			return ColumnNames;
		
	  }
	  
	  /*SAPGUITextFieldSetOnExistence
	   * Objective - To set the text field only if it exist  
	   * parameter -  lbl: Control Name , txtvalue - value to be set, LabelClassName - Class name of the field.
	   * created by Venkata Siva kumar     
	   */
	  public void SAPGUITextFieldSetOnExistence(String lbl, String txtValue, String LabelClassName) 
	  {
		   try 
		   {
			   arg[0] = new Variant(lbl);
			   arg[1] = new Variant(LabelClassName);
			   session = new ActiveXComponent(Obj.invoke("FindByName", arg).toDispatch());
			   session.setProperty("text", txtValue);
		   }
			catch (Exception e) 
			{
				return;
				 // there is no need to handle this
			}
			 			
		}
	  
	  /*SAPGUIButtonClickOnExist
	   * Objective - To click on button if found   
	   * parameter -  lbl: Control Name  
	   * returns true if the button found and clicked else returns false 
	   * created by Venkata Siva kumar     
	   */
	  public boolean SAPGUIButtonClickOnExist(String lblname)
	  {
		  try 
		  {
			  SAPGUIClick(lblname, SAPGuiClassName.GuiButtonClassName);
			  return true;
		  }
		  catch (Exception e) 
		  {
			  return false;
			 // there is no need to handle this
		  }
	  }
	  
	  /*SAPGUISelectColumnInTableControl
	   * Objective - to select/de select the column the table control
	   * parameter  : Element Name - name of the element , column name to be selected or delselected 
	   * parameter : selectOrDeselect - true for selecting the column and false for De-selecting the column
	   * created by Venkata Siva kumar    
	  */
	public void SAPGUISelectColumnInTableControl(String Lbl, String ColumName , boolean selectOrDeselect) 
	{ 		
		
		ActiveXComponent currentObj;
		arg[0] = new Variant(Lbl);
		arg[1] = new Variant(SAPGuiClassName.GuiTableControlClassName);				 
		session = new ActiveXComponent(getSession().invoke("FindByName",  arg).toDispatch());	
		Dispatch dis = Dispatch.call(session, "Columns").toDispatch();
		int count = Dispatch.call(dis, "count").toInt(); 			 
		Dispatch[] arr = new Dispatch[count];
		for(int i=0;i<count;i++)
		{
			currentObj = new ActiveXComponent(session.invoke("Columns", i).toDispatch());
			if( currentObj.getProperty("title").toString().trim().equals(ColumName))
			{
				currentObj.setProperty("selected", selectOrDeselect);
				return ;
			}
		}

	}
	
	/* SAPGUITableControlGetRowIndexByItsValue
	 * Objective - Function searches the cell value in the table and returns the row number
	 * Parameter  lbl - Name of the TableControl, ColumnNmae - Name of the column, Cell value - Value of the cell text to be mathced 
	 * returns -1 if the  value is not matched with any cell, else it returns the  row number of the first occurance 
	 * Created by Venkata Siva kumar 
	 */	
	@SuppressWarnings("deprecation")
	public int SAPGUITableControlGetRowIndexByItsValue(String lbl, String CellValue, String ColumnName ) 
	{
		String str = null;
		int RowCount = SAPGUITableControlGetVisibleRowCount(lbl);
		for(int i=0 ; i<RowCount; i++)
		{		
			str = SAPGUITableControlGetCellObject(lbl,i,ColumnName).getProperty("text").toString().trim();
			if(str.equals(CellValue) )
				return i+1;
		}
		return -1 ;
	
	} 
	
	/* SAPGUITableControlSelectORDeselectAllColumns
	 * Objective - Function to select / deslect all the columns in table 
	 * Parameter  lbl - Name of the TableControl, SelectOrDeselect -  Send true to select all columns and false to deselect all columns
	 * Created by Venkata Siva kumar 
	 */	
	@SuppressWarnings("deprecation")
	public void SAPGUITableControlSelectORDeselectAllColumns(String lbl, boolean SelectOrDeselect ) 
		{
		
			String str = null;
			ActiveXComponent currentObj;
			arg[0] = new Variant(lbl);
			arg[1] = new Variant(SAPGuiClassName.GuiTableControlClassName);				 
			session = new ActiveXComponent(getSession().invoke("FindByName",  arg).toDispatch());	
			if (SelectOrDeselect) 
				session.invoke("SelectAllColumns");
			else
				session.invoke("DeselectAllColumns");
		}	
	
	/* SAPGUITableControlInsertTextinAllRows 
	 * Objective - Function sets all the visible rows of a particular column 
	 * Parameter  lbl - Name of the TableControl, ColumnName - Name of the column, 
	 * Parameter ValueToBeSet - Value to be set in all the rows of the column	 
	 * Created by Venkata Siva kumar 
	 */
	@SuppressWarnings("deprecation")
	public void SAPGUITableControlInsertTextinAllRows(String lbl, String ColumName, String ValueToBeSet ) 
	{
			String str = null;
			ActiveXComponent currentObj;
			int RowCount = SAPGUITableControlGetVisibleRowCount(lbl);
			for(int i=0 ; i<RowCount; i++)
			{		
				currentObj = SAPGUITableControlGetCellObject(lbl,i,ColumName);
				currentObj.setProperty("text", ValueToBeSet  );	         				
			}	
	}
	
	  
	  /*SAPGuiGridClickOnCell
	   * Objective -To click on cell in Grid
	   * parameter  : partial id string , row number and column name 
	   * created by Venkata Siva kumar    
	   */
	  public void SAPGuiGridClickOnCell(String idstr, int row, String column) throws Exception {
		 
			  String id = getSAPObjectIDHelperMethod(getSession(), "ID", idstr, "", "");
			  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
			  Dispatch.call(session, "setCurrentCell", new Object[] { row, column });
			  Dispatch.call(session, "clickCurrentCell");
	  }
	  
	  /*SAPGuiGridSelectRow
	   * Objective - to  select the row in the grid 
	   * parameter  : partial id string , row number and column name 
	   * created by Venkata Siva kumar     
	   */
	  public void SAPGuiGridSelectRow(String idstr, int row, String column) throws Exception {
		 
		  String id = getSAPObjectIDHelperMethod(getSession(), "ID", idstr, "", "");
		  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());	 
		  session.setProperty("selectedRows", row);
	  }
	  
	  /*SAPGuiGridSelectOrDeselectColumn
	   * Objective - to  select or deselect the column  in the grid 
	   * parameter  : partial id string , column name , SelectOrDeselect - Pass true to select column and false to Deselect the column
	   * created by Venkata Siva kumar    
	   */
	  public void SAPGuiGridSelectOrDeselectColumn(String idstr, String column, boolean SelectOrDeselect) throws Exception
	  {
		 
		  String id = getSAPObjectIDHelperMethod(getSession(), "ID", idstr, "", "");
		  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());	 
		  if (SelectOrDeselect)
			  Dispatch.call(session, "selectColumn", column);
		 
		  else
			  Dispatch.call(session, "deselectColumn", column);
		 
	  }
	  
	  /*SAPGuiGridSelectOrDeselectColumn
	   * Objective - to select entire grid
	   * parameter  : partial id string 
	   * created by Venkata Siva kumar     
	   */
	  public void SAPGuiGridSelectGrid(String idstr) throws Exception {
		 
		  String id = getSAPObjectIDHelperMethod(getSession(), "ID", idstr, "", "");
		  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());	 
		  Dispatch.call(session, "SelectAll");	 
		 
		 
	  }
	  
	  /*SAPGUIGridGetColumnCount
	   * Objective - to fetch the grid  column count
	   * parameter  : partial id string 
	   * returns - number of columns in grid.
	   * created by Venkata Siva kumar     
	   */
	  public int SAPGUIGridGetColumnCount(String idstr) throws Exception {
		 
		  String id = getSAPObjectIDHelperMethod(getSession(), "ID", idstr, "", "");
		  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());	 
		  Variant var =  session.getProperty("ColumnCount");
		 return  var.toInt();		 
		 
	  }
	  
	  /*SAPGuiSetFocusOnEdit
	   * Objective -Set focus on Gui Edit text box
	   * parameter  lbl : Name of the object  
	   * created by Venkata Siva kumar     
	   */	  
	  public  void SAPGuiSetFocusOnEdit(String lbl) throws Exception	  
	  {
			arg[0] = new Variant(lbl);
			arg[1] = new Variant("GuiTextClassName");
			session = new ActiveXComponent(Obj.invoke("FindByName", arg).toDispatch());
			session.invoke("setFocus");			 
	  }
	  
	  /*SAPGuiSetFocusOnEdit
	   * Objective -Sets the value in specific cell, (*cell should be editable) 
	   * parameter  lbl : Name of the object  
	   * Column Name : Name of the column , Row - Row number 
	   * created by Venkata Siva kumar     
	   */	  
	  public  void SAPGuiGridSetTextinCell(String idstr, int Row, String ColumnName, String ValueToBeSet) throws Exception {
		 
		  String id = getSAPObjectIDHelperMethod(getSession(), "ID", idstr, "", "");
		  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());	 		 
		  Dispatch.call(session, "ModifyCell", Row,  ColumnName,ValueToBeSet );
		
	  }
	  
	  /*SAPTableColumnTextValueSet
	   * Objective - To enter corresponding value to the column name in a table (eg - MARA)
	   * parameter  :colPartialStr, columnText, textFieldLbl,textType,value
	   * created by Venkata Siva Kumar    
	   */
	  public void SAPTableColumnTextValueSet(String colPartialStr, String columnText, String textPartialStr,String textType,String value) throws Exception{
		  String id = getSAPObjectIDHelperMethod(getSession(), "Name", colPartialStr, "Text", columnText);
		  String array[] = id.split(",");
		  String idText = getSAPObjectIDHelperMethod(getSession(), "Name", textPartialStr, "Type", textType);
		  String textArray[] = idText.split(",");
		  String finalStr = textArray[0]+","+array[1];
		  session = new ActiveXComponent(getSession().invoke("FindById",finalStr).toDispatch());
		  session.invoke("setFocus");
		  session.setProperty("text", value);
		  	    
	  }

	  /*SAPStatusBarGetText
	   * Objective - To get the text from status bar
	   * parameter  : type and name
	   * created by Venkata Siva Kumar    
	   */
	  public String SAPStatusBarGetText(String type,String name) throws Exception {

			String id = getSAPObjectIDHelperMethod(getSession(), "Name", name, "Type", type);
			session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
			String text =  session.getProperty("Text").toString().trim() ;
			return text;	
		  
	  }
	  
	  /*
		 * SAPGUICheckBoxOnOffUsingPartialText Objective, to select or unselect the check box
		 * parameter lbl : partial id, Name of the check box parameter OnOrOff : Pass true to select
		 * the radio button and false to deselect the radio button 
		 * created by Venkata Siva Kumar
		 * 
		 */
		public void SAPGUICheckBoxOnOffUsingPartialText(String partialId, String lbl, Boolean OnOrOff) {
			ActiveXComponent wnd0 = getSession();
			String id = getSAPObjectIDHelperMethod(getSession(), "ID", partialId, "Type", "GuiCheckBox");
			  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());
			if (OnOrOff)
			{
				session.invoke("setFocus");
				session.setProperty("selected", -1);
			}
			else
			{
				session.invoke("setFocus");
				session.setProperty("selected", 0);
			}

			setSession(wnd0);
			session = wnd0;
		}
		
		/*
		 * SAPGUICheckBoxWithinGuiShell Objective, to select or unselect the check box
		 * parameter lbl : partial id of shell,row name,column name, true or false to select or deselect check box
		 * created by Venkata Siva Kumar
		 * 
		 */
		public void SAPGUICheckBoxWithinGuiShell(int k, String colName, Boolean OnOrOff) {
			
			String id = getSAPObjectIDHelperMethod(getSession(), "Name", "shell", "Type", "GuiShell");
			  session = new ActiveXComponent(getSession().invoke("FindById",id).toDispatch());

			if (OnOrOff)
			{
				session.setProperty("currentCellRow", k);
				session.invoke("setFocus");
				Dispatch.call(session,"modifyCheckbox",k,colName,-1);
				Dispatch.call(session,"triggerModified");
			}
			else
			{
				session.setProperty("currentCellRow", k);
				session.invoke("setFocus");
				Dispatch.call(session,"modifyCheckbox",k,colName,0);
				Dispatch.call(session,"triggerModified");
			}

		}
		
		/*SAPPopUpMultiTextValueSet
		   * Objective - To enter multiple value in the pop up text fields [eg - Enter multiple articles/request IDs in push button pop up window]
		   * parameter  :textFieldLbl,textType,rownum,value
		   * created by Venkata Siva Kumar    
		   */
		  public void SAPPopUpMultiTextValueSet(String textPartialStr,String textType,int row,String value,int window) throws Exception{
			  ActiveXComponent wnd0 = getSession();
			  setSession(new ActiveXComponent(getParentSession().invoke("FindById", "wnd["+window+"]").toDispatch()));
			  String idText = getSAPObjectIDHelperMethod(getSession(), "Name", textPartialStr, "Type", textType);
			  String textArray[] = idText.split(",");
			  String finalStr = textArray[0]+","+row+"]";
			  session = new ActiveXComponent(getSession().invoke("FindById",finalStr).toDispatch());
			  session.invoke("setFocus");
			  session.setProperty("text", value);
			  setSession(wnd0); 
		  }
}