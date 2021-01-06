package com.automationanywhere.botcommand.samples.commands;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;

import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;

@BotCommand
@CommandPkg(label = "[[CreateDocument.label]]", description = "[[CreateDocument.description]]", icon = "sample.svg", name = "CreateDocument")

public class CreateDocument {

    //Messages read from full qualified property file name and provide i18n capability.
    private static final Messages MESSAGES = MessagesFactory
            .getMessages("com.automationanywhere.botcommand.samples.messages");

    @Execute
    public void execute (

            @Idx(index = "1", type = TEXT)
            @Pkg(label = "[[CreateDocument.filePath.label]]")
            @NotEmpty String filePath,

            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[CreateDocument.fileName.label]]")
            @NotEmpty String fileName,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "[[CreateDocument.createParagraph.label]]")
            String paragraph

    ) throws Exception

    {
            // Create document variable
            XWPFDocument document = new XWPFDocument();
            //Create Paragraph Variable
            XWPFParagraph createParagraph = document.createParagraph();
            //Create run variable to Run/Execute Methods
            XWPFRun run = createParagraph.createRun();
            //Call createDocument method to create a new Word document
            createDocument(document, run, paragraph, filePath, fileName);

    }

    private static void createDocument (XWPFDocument document, XWPFRun run, String paragraph, String filePath, String fileName) throws Exception {
        if (paragraph != "") {
            run.setText(paragraph);
        }
        FileOutputStream fileOutputStream = new FileOutputStream(new File(filePath+"\\"+ fileName+".docx"));
        document.write(fileOutputStream);
        document.close();
        fileOutputStream.close();
    }

}
