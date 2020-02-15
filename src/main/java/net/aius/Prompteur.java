package net.aius;

import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class Prompteur {

    public static void main(String[] args) {
        List<String> categories = Arrays.asList(args);

        Prompteur prompteur = new Prompteur();
        prompteur.createPPT(categories);
    }

    private void createPPT(List<String> categories) {
        XMLSlideShow ppt = new XMLSlideShow();

        for (String str : categories) {
            createSlide(ppt, str);
        }

        exportPPT(ppt);
    }

    private void createSlide(XMLSlideShow ppt, String str) {
        XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);

        XSLFSlideLayout layout = defaultMaster.getLayout(SlideLayout.TITLE_ONLY);
        XSLFSlide slide = ppt.createSlide(layout);

        XSLFTextShape shape = slide.getPlaceholder(0);
        shape.clearText();

        XSLFTextParagraph paragraph = shape.addNewTextParagraph();
        XSLFTextRun run = paragraph.addNewTextRun();
        run.setText(str);
        run.setFontSize(80.D);

        paragraph.addLineBreak();
        Dimension dimension = ppt.getPageSize();
        shape.setAnchor(new Rectangle(
                0,
                (int) (dimension.getHeight() / 2 - shape.getTextHeight() / 3),
                (int) ppt.getPageSize().getWidth(),
                (int) shape.getTextHeight()
        ));
    }

    private void exportPPT(XMLSlideShow ppt) {
        try {
            FileOutputStream out = new FileOutputStream("prompteur.pptx");
            ppt.write(out);
            out.close();

            System.out.println("Votre fichier \"prompteur.pptx\" est maintenant disponible.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
