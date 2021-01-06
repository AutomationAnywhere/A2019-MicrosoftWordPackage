package com.automationanywhere.botcommand.samples.commands;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import static com.automationanywhere.commandsdk.model.AttributeType.FILE;
import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;

@BotCommand
@CommandPkg(label = "[[Bookmark.label]]", description = "[[Bookmark.description]]", icon = "sample.svg", name = "Bookmark")


public class InsertTextAtBookmark {

    //Messages read from full qualified property file name and provide i18n capability.
    private static final Messages MESSAGES = MessagesFactory
            .getMessages("com.automationanywhere.botcommand.samples.messages");

    @Execute
    public void execute (

            @Idx(index = "1", type = FILE)
            @Pkg(label = "[[Bookmark.filePath.label]]")
            @NotEmpty String filePath,

            @Idx(index = "2", type = TEXT)
            @Pkg(label = "[[Bookmark.bookmarkName.label]]")
            @NotEmpty String bookmarkName,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "[[Bookmark.insertText.label]]")
            @NotEmpty String insertText

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
        insertText(document, bookmarkName, insertText, filePath);

    }

    private static void insertText (XWPFDocument document, String bookmarkName, String insertText, String filePath) throws Exception {

        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for(XWPFParagraph paragraph : paragraphs) {
            CTP ctp = paragraph.getCTP();
            // Get all bookmarks and loop through them
            List<CTBookmark> bookmarks = ctp.getBookmarkStartList();
            for(CTBookmark bookmark : bookmarks) {
                // Extract the name of the bookmark
//                String bookmarkNames = bookmark.getName();
//                System.out.println(bookmarkName);
                if (bookmark.getName().equals(bookmarkName)) {
                    // Create a new RSID (revision identifier) and add text
                    CTR ctr = ctp.addNewR();
                    CTText text = ctr.addNewT();
                    text.setStringValue(insertText);
                }
//                else if (bookmark.getName().equals("myBookmark")) {
//                    CTR ctr = ctp.addNewR();
//                    CTText text = ctr.addNewT();
//                    text.setStringValue("1st Bookmark");
//                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        document.write(fileOutputStream);
        document.close();
        fileOutputStream.close();
    }

}
