package com.open.source;
import com.aspose.slides.*;

import java.awt.Color;

public class AcceleratorPackage {
    public static void main(String[] args) {
        // Create a new presentation
        Presentation pres = new Presentation();

        // Get the first slide
        ISlide slide = pres.getSlides().get_Item(0);

        // Add first rectangle (Block 1)
        IAutoShape block1 = slide.getShapes().addAutoShape(ShapeType.Rectangle, 100, 100, 150, 80);
        block1.getTextFrame().setText("System A");
        block1.getFillFormat().setFillType(FillType.Solid);
        //block1.getFillFormat().getSolidFillColor().setColor(Color.ORANGE);

        // Add second rectangle (Block 2)
        IAutoShape block2 = slide.getShapes().addAutoShape(ShapeType.Rectangle, 400, 100, 150, 80);
        block2.getTextFrame().setText("System B");
        block2.getFillFormat().setFillType(FillType.Solid);
        //block2.getFillFormat().getSolidFillColor().setColor(Color.CYAN);

        // Add a connector between Block 1 and Block 2
        IConnector connector = slide.getShapes().addConnector(ShapeType.BentConnector2, 0, 0, 10, 10);
        connector.setStartShapeConnectedTo(block1);
        connector.setEndShapeConnectedTo(block2);
        connector.getLineFormat().getFillFormat().setFillType(FillType.Solid);
        connector.getLineFormat().getFillFormat().getSolidFillColor().setColor(Color.BLACK);
        connector.getLineFormat().setWidth(2);

        // Save the presentation
        pres.save("inventory_report.pptx", SaveFormat.Pptx);
    }
}
