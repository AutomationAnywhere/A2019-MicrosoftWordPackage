package com.automationanywhere.botcommand.samples.commands;


import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import static com.automationanywhere.commandsdk.model.AttributeType.FILE;
import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;

@BotCommand
@CommandPkg(label = "[[ReplaceText.label]]", description = "[[ReplaceText.description]]", icon = "sample.svg", name = "ReplaceText")

public class ReplaceText {

    //Messages read from full qualified property file name and provide i18n capability.
    private static final Messages MESSAGES = MessagesFactory
            .getMessages("com.automationanywhere.botcommand.samples.messages");

    @Execute
    public void execute (

            @Idx(index = "1", type = FILE)
            @Pkg(label = "[[ReplaceText.filePath.label]]")
            @NotEmpty String filePath,

            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[ReplaceText.replaceText.label]]")
            @NotEmpty String replaceText,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "[[ReplaceText.newText.label]]")
            @NotEmpty String newText

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
        //Call createDocument method to create a new Word document
        replaceText(document, replaceText, newText, filePath);

    }

    private static void replaceText (XWPFDocument document, String replaceText, String newText, String filePath) throws Exception {
        for (XWPFParagraph getParagraph : document.getParagraphs()) {
            List<XWPFRun> getRuns = getParagraph.getRuns();
            if (getRuns != null) {
                for (XWPFRun run : getRuns) {
                    String text = run.getText(0);
                    if (text != null && text.contains(replaceText)) {
                        text = text.replace(replaceText, newText);
                        run.setText(text, 0);
                    }
                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        document.write(fileOutputStream);
        document.close();
        fileOutputStream.close();
    }


}
