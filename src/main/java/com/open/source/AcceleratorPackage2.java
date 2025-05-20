package com.open.source;

import com.aspose.slides.*;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.open.source.model.Edge;
import com.open.source.model.Node;

import java.awt.*;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class AcceleratorPackage2 {

    public static void main(String[] args) throws Exception {
        // Load JSON file
        File jsonFile = new File("C:\\Personal\\coding\\java\\ms_poc\\aspose-slides-poc\\lineage.json"); // Replace with your path
        ObjectMapper mapper = new ObjectMapper();
        JsonNode root = mapper.readTree(jsonFile);

        // Parse nodes and edges
        List<Node> nodes = new ArrayList<Node>();
        Map<String, Node> nodeMap = new HashMap<>();
        for (JsonNode node : root.get("nodes")) {
            Node n = new Node();
            n.id = node.get("id").asText();
            n.name = node.get("name").asText();
            n.type = node.get("type").asText();
            nodes.add(n);
            nodeMap.put(n.id, n);
        }

        List<Edge> edges = new ArrayList<>();
        for (JsonNode edge : root.get("edges")) {
            Edge e = new Edge();
            e.from = edge.get("from").asText();
            e.to = edge.get("to").asText();
            e.relation = edge.get("relation").asText();
            edges.add(e);
        }

        // Create PPT
        // Create PPT
        Presentation pres = new Presentation();
        ISlide slide = pres.getSlides().get_Item(0);

        float slideWidth = (float) pres.getSlideSize().getSize().getWidth();

        // Keep track of shapes and their positions
        Map<String, IShape> shapeMap = new HashMap<>();
        int startX = 100;
        int startY = 150;
        int xStep = 300;
        int yStep = 150;

        // Layout shapes horizontally
        int index = 0;
        for (Node n : nodes) {
            float x = startX + (index * xStep);
            float y = startY;

            int shapeType = n.type.equalsIgnoreCase("database") ? ShapeType.Ellipse : ShapeType.Rectangle;

            IAutoShape shape = slide.getShapes().addAutoShape(shapeType, x, y, 150, 80);
            shape.getTextFrame().setText(n.name);
            shape.getFillFormat().setFillType(FillType.Solid);
            shape.getFillFormat().getSolidFillColor().setColor(n.type.equalsIgnoreCase("database") ? Color.LIGHT_GRAY : Color.ORANGE);
            shape.getLineFormat().getFillFormat().getSolidFillColor().setColor(Color.BLACK);

            shapeMap.put(n.id, shape);
            index++;
        }

        // Draw connectors and add labels
        for (Edge e : edges) {
            IShape fromShape = shapeMap.get(e.from);
            IShape toShape = shapeMap.get(e.to);
            if (fromShape == null || toShape == null) continue;

            // Draw straight connector
            IConnector connector = slide.getShapes().addConnector(ShapeType.Line, 0, 0, 10, 10);
            connector.setStartShapeConnectedTo(fromShape);
            connector.setEndShapeConnectedTo(toShape);
            connector.getLineFormat().getFillFormat().setFillType(FillType.Solid);
            connector.getLineFormat().getFillFormat().getSolidFillColor().setColor(Color.BLACK);
            connector.getLineFormat().setWidth(2);

            // Calculate label box position between from and to
            float labelX = (fromShape.getX() + toShape.getX()) / 2;
            float labelY = (fromShape.getY() + toShape.getY()) / 2 - 20;

            // Estimate label width
            int charWidth = 7;
            int maxTextWidth = 100;
            String labelText = e.relation;
            if (labelText.length() * charWidth > maxTextWidth) {
                int maxChars = maxTextWidth / charWidth;
                labelText = labelText.substring(0, Math.min(maxChars, labelText.length())) + "...";
            }

            IAutoShape label = slide.getShapes().addAutoShape(ShapeType.Rectangle, labelX, labelY, maxTextWidth, 25);
            label.getTextFrame().setText(labelText);
            label.getFillFormat().setFillType(FillType.NoFill);
            label.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
        }

        // Save presentation
        pres.save("swimlane_diagram_fixed.pptx", SaveFormat.Pptx);
        System.out.println("PPT generated: swimlane_diagram_fixed.pptx");

    }
}
