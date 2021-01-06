/*
 * Copyright (c) 2020 Automation Anywhere.
 * All rights reserved.
 *
 * This software is the proprietary information of Automation Anywhere.
 * You shall use it only in accordance with the terms of the license agreement
 * you entered into with Automation Anywhere.
 */
package com.automationanywhere.botcommand.samples.commands.ui;

import static com.automationanywhere.commandsdk.model.AttributeType.*;

import java.time.ZonedDateTime;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.TableValue;
import com.automationanywhere.botcommand.data.impl.WindowValue;
import com.automationanywhere.botcommand.data.model.Point;
import com.automationanywhere.botcommand.data.model.Taskbot;
import com.automationanywhere.botcommand.data.model.UIObject;
import com.automationanywhere.botcommand.data.model.image.Image;
import com.automationanywhere.botcommand.data.model.property.Property;
import com.automationanywhere.botcommand.data.model.record.Record;
import com.automationanywhere.botcommand.data.model.region.Region;
import com.automationanywhere.commandsdk.annotations.BotCommand;
import com.automationanywhere.commandsdk.annotations.CommandPkg;
import com.automationanywhere.commandsdk.annotations.Execute;
import com.automationanywhere.commandsdk.annotations.Idx;
import com.automationanywhere.commandsdk.annotations.Pkg;
import com.automationanywhere.commandsdk.annotations.rules.CodeType;
import com.automationanywhere.commandsdk.annotations.rules.VariableType;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.automationanywhere.core.security.SecureString;

/**
 * @author Raj Singh Sisodia
 *
 */
//@BotCommand
//@CommandPkg(label = "UI Demo", description = "Demonstrates the provided UI elements", name = "uiDemo")

public class SampleUI {

    @Idx(index = "1", type = TEXT)
    @Pkg(label = "Text type")
    String text;
	
    @Idx(index = "2", type = NUMBER)
    @Pkg(label = "Number type")
    String num;
    
    @Idx(index = "3", type = BOOLEAN)
    @Pkg(label = "Boolean type", description = "A check box also supports children.")
    Boolean bool;
    
    @Idx(index = "3.1", type = TEXT)
    @Pkg(label = "Text type", description = "Child of boolean, this gets evaluated only when the checkbox is true")
    String boolChild1;
    
    @Idx(index = "4", type = SELECT, options = {
    		@Idx.Option(index ="4.1", pkg = @Pkg(label = "Select Option 1", value = "Select option 1"))
    })
    @Pkg(label = "Select type")
    String select;
    
    @Idx(index = "5", type = RADIO, options = {
    		@Idx.Option(index ="5.1", pkg = @Pkg(label = "Select Option 1", value = "Radio option 1")),
    		@Idx.Option(index ="5.2", pkg = @Pkg(label = "Select Option 1", value = "Radio option 2"))
    })
    @Pkg(label = "Radio type")
    String radio;
    
    @Idx(index = "6", type = FILE)
    @Pkg(label = "File type")
    String file;
    
    @Idx(index = "7", type = AttributeType.VARIABLE)
    @Pkg(label = "Variable type", description = "Variable accepting only Boolean types")
    @VariableType(DataType.BOOLEAN)
    Boolean variable;
    
    @Idx(index = "9", type = TEXT)
    @Pkg(label = "TextArea type")
    String textarea;
    
    @Idx(index = "10", type = CODE)
    @Pkg(label = "Code type", description = "Accepting javascript")
    @CodeType(value = "text/javascript")
    String code;
    
    @Idx(index = "11", type = DICTIONARY)
    @Pkg(label = "Dictionary type")
    Map<String, String> dictionaryType;
    
    @Idx(index = "12", type = LIST)
    @Pkg(label = "List type")
    List<String> listType;
    
    @Idx(index = "13", type = BOOLEAN)
    @Pkg(label = "Boolean type")
    String boolType;
    
    @Idx(index = "14", type = TASKBOT)
    @Pkg(label = "TaskBot type")
    Taskbot taskBot;
    
    @Idx(index = "15", type = WINDOW)
    @Pkg(label = "Window type")
    WindowValue window;
    
    @Idx(index = "16", type = TABLE)
    @Pkg(label = "Table type")
    TableValue tableType;
    
    @Idx(index = "17", type = CREDENTIAL)
    @Pkg(label = "Credential type")
    SecureString credential;
    
    @Idx(index = "18", type = KEYPRESS)
    @Pkg(label = "Keypress type")
    String keyPress;
    
    @Idx(index = "19", type = DATETIME)
    @Pkg(label = "DateTime type")
    ZonedDateTime dateTime;
    
    @Idx(index = "20", type = VARIABLEMAP)
    @Pkg(label = "VariableMap type")
    Map<String, Value> variableMap;
    
    @Idx(index = "21", type = COORDINATE)
    @Pkg(label = "Coordinate type")
    Point coordinate;
    
    @Idx(index = "22", type = IMAGE)
    @Pkg(label = "Text type")
    Image image;

    @Idx(index = "23", type = REGION)
    @Pkg(label = "Text type")
    Region region;
    
    @Idx(index = "24", type = PROPERTIES)
    @Pkg(label = "Properties type")
    Set<Property> properties;
    
    @Idx(index = "25", type = UIOBJECT)
    @Pkg(label = "UIObject type")
    UIObject uiObject;
    
    @Idx(index = "26", type = RECORD)
    @Pkg(label = "Record type")
    Record record;
    
    @Idx(index = "27", type = HELP)
    @Pkg(label = "Help type", description = "Help description")
    Record help;
    
    
    @Execute
    public void execute() {
    	
    }

	public String getText() {
		return text;
	}

	public void setText(String text) {
		this.text = text;
	}

	public String getNum() {
		return num;
	}

	public void setNum(String num) {
		this.num = num;
	}

	public Boolean getBool() {
		return bool;
	}

	public void setBool(Boolean bool) {
		this.bool = bool;
	}

	public String getBoolChild1() {
		return boolChild1;
	}

	public void setBoolChild1(String boolChild1) {
		this.boolChild1 = boolChild1;
	}

	public String getSelect() {
		return select;
	}

	public void setSelect(String select) {
		this.select = select;
	}

	public String getRadio() {
		return radio;
	}

	public void setRadio(String radio) {
		this.radio = radio;
	}

	public String getFile() {
		return file;
	}

	public void setFile(String file) {
		this.file = file;
	}

	public Boolean getVariable() {
		return variable;
	}

	public void setVariable(Boolean variable) {
		this.variable = variable;
	}

	public String getTextarea() {
		return textarea;
	}

	public void setTextarea(String textarea) {
		this.textarea = textarea;
	}

	public String getCode() {
		return code;
	}

	public void setCode(String code) {
		this.code = code;
	}

	public Map<String, String> getDictionaryType() {
		return dictionaryType;
	}

	public void setDictionaryType(Map<String, String> dictionaryType) {
		this.dictionaryType = dictionaryType;
	}

	public List<String> getListType() {
		return listType;
	}

	public void setListType(List<String> listType) {
		this.listType = listType;
	}

	public String getBoolType() {
		return boolType;
	}

	public void setBoolType(String boolType) {
		this.boolType = boolType;
	}

	public Taskbot getTaskBot() {
		return taskBot;
	}

	public void setTaskBot(Taskbot taskBot) {
		this.taskBot = taskBot;
	}

	public Record getHelp() {
		return help;
	}

	public void setHelp(Record help) {
		this.help = help;
	}

	public WindowValue getWindow() {
		return window;
	}

	public void setWindow(WindowValue window) {
		this.window = window;
	}

	public TableValue getTableType() {
		return tableType;
	}

	public void setTableType(TableValue tableType) {
		this.tableType = tableType;
	}

	public SecureString getCredential() {
		return credential;
	}

	public void setCredential(SecureString credential) {
		this.credential = credential;
	}

	public String getKeyPress() {
		return keyPress;
	}

	public void setKeyPress(String keyPress) {
		this.keyPress = keyPress;
	}

	public ZonedDateTime getDateTime() {
		return dateTime;
	}

	public void setDateTime(ZonedDateTime dateTime) {
		this.dateTime = dateTime;
	}

	public Map<String, Value> getVariableMap() {
		return variableMap;
	}

	public void setVariableMap(Map<String, Value> variableMap) {
		this.variableMap = variableMap;
	}

	public Point getCoordinate() {
		return coordinate;
	}

	public void setCoordinate(Point coordinate) {
		this.coordinate = coordinate;
	}

	public Image getImage() {
		return image;
	}

	public void setImage(Image image) {
		this.image = image;
	}

	public Region getRegion() {
		return region;
	}

	public void setRegion(Region region) {
		this.region = region;
	}

	public Set<Property> getProperties() {
		return properties;
	}

	public void setProperties(Set<Property> properties) {
		this.properties = properties;
	}

	public UIObject getUiObject() {
		return uiObject;
	}

	public void setUiObject(UIObject uiObject) {
		this.uiObject = uiObject;
	}

	public Record getRecord() {
		return record;
	}

	public void setRecord(Record record) {
		this.record = record;
	}
    
    
}
