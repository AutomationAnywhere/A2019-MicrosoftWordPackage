package com.automationanywhere.botcommand.samples.commands;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import static com.automationanywhere.commandsdk.model.AttributeType.FILE;
import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;

@BotCommand
@CommandPkg(label = "[[AddParagraph.label]]", description = "[[AddParagraph.description]]", icon = "sample.svg", name = "AddParagraph")


public class AddParagraph {

    //Messages read from full qualified property file name and provide i18n capability.
    private static final Messages MESSAGES = MessagesFactory
            .getMessages("com.automationanywhere.botcommand.samples.messages");

    @Execute
    public void execute (

            @Idx(index = "1", type = FILE)
            @Pkg(label = "[[AddParagraph.filePath.label]]")
            @NotEmpty String filePath,

            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[AddParagraph.addParagraph.label]]")
            @NotEmpty String paragraph

    ) throws Exception


    {
        //
        FileInputStream fileInputStream = new FileInputStream(filePath);
        // Create document variable
        XWPFDocument document = new XWPFDocument(fileInputStream);
        //Create Paragraph Variable
        XWPFParagraph createParagraph = document.createParagraph();
        //Create run variable to Run/Execute Methods
        XWPFRun run = createParagraph.createRun();
        // Create cursor variable to insert new paragraph
        XmlCursor cursor = createParagraph.getCTP().newCursor();
        //Call createDocument method to create a new Word document
        addParagraph(document, run, paragraph, filePath);



    }

    private static void addParagraph (XWPFDocument document, XWPFRun run, String paragraph, String filePath) throws Exception {
        run.setText(paragraph);
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        document.write(fileOutputStream);
        document.close();
        fileOutputStream.close();
    }

}
