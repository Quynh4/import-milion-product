package com.milionrowsexcelfile;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

public class ProductSheetHandler extends DefaultHandler {

    private SharedStringsTable sst;
    private ProductRepository productRepository;

    private String lastContents;
    private boolean nextIsString;
    private List<Product> batch = new ArrayList<>();
    private int batchSize = 1000;

    private int col = 0;
    private String code, name;
    private double price;

    public ProductSheetHandler(SharedStringsTable sst, ProductRepository repo) {
        this.sst = sst;
        this.productRepository = repo;
    }

    public void startElement(String uri, String localName, String qName, Attributes attributes) {
        if ("c".equals(qName)) {
            String cellType = attributes.getValue("t");
            nextIsString = "s".equals(cellType);
        }
    }

    public void endElement(String uri, String localName, String qName) {
        if ("v".equals(qName)) {
            String value = nextIsString ? sst.getItemAt(Integer.parseInt(lastContents)).getString() : lastContents;

            switch (col) {
                case 0 -> code = value;
                case 1 -> name = value;
                case 2 -> price = Double.parseDouble(value);
            }
            col++;
        } else if ("row".equals(qName)) {
            // Nếu có đủ 3 cột thì lưu
            if (col == 3) {
                batch.add(new Product(null, code, name, price));
                if (batch.size() >= batchSize) {
                    productRepository.saveAll(batch);
                    batch.clear();
                }
            }
            col = 0; // reset dòng mới
        }
    }

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
