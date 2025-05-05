package com.milionrowsexcelfile;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

public class ProductSheetHandler extends DefaultHandler {

    private final SharedStringsTable sst;
    private final ProductRepository productRepository;

    private String lastContents;
    private boolean nextIsString;

    private final List<Product> batch = new ArrayList<>();
    private final int batchSize = 100000;

    private int col = 0;
    private String name;
    private double price;
    private int quantity;
    private boolean isFirstRow = true;

    public ProductSheetHandler(SharedStringsTable sst, ProductRepository repo) {
        this.sst = sst;
        this.productRepository = repo;
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {
        if ("c".equals(qName)) {
            String cellType = attributes.getValue("t");
            nextIsString = "s".equals(cellType);
        }
    }

    @Override
    public void endElement(String uri, String localName, String qName) {
        if ("v".equals(qName)) {
            String value = nextIsString
                    ? sst.getItemAt(Integer.parseInt(lastContents)).getString()
                    : lastContents;

            // Bắt đầu từ cột 1 (bỏ qua cột 0)
            switch (col) {
                case 1 -> name = value;
                case 2 -> price = Double.parseDouble(value);
                case 3 -> quantity = Integer.parseInt(value);
            }
            col++;
        } else if ("row".equals(qName)) {
            if (isFirstRow) {
                isFirstRow = false; // bỏ qua dòng tiêu đề
            } else if (col >= 4) {
                Product product = new Product();
                product.setName(name);
                product.setPrice(price);
                product.setQuantity(quantity);
                batch.add(product);
            }

            if (batch.size() >= batchSize) {
                productRepository.saveAll(batch);
                batch.clear();
            }

            col = 0; // reset cột cho dòng mới
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        lastContents = new String(ch, start, length);
    }

    @Override
    public void endDocument() {
        if (!batch.isEmpty()) {
            productRepository.saveAll(batch);
        }
    }
}
