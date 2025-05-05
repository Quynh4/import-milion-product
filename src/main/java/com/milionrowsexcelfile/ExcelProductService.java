package com.milionrowsexcelfile;

import com.milionrowsexcelfile.Product;
import com.milionrowsexcelfile.ProductRepository;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.*;
import java.util.concurrent.*;

@Service
public class ExcelProductService {

    private final ProductRepository productRepository;
    private static final int THREAD_COUNT = 4; // Số luồng song song


    public ExcelProductService(ProductRepository productRepository) {
        this.productRepository = productRepository;
    }

    public void importProducts(MultipartFile file) {
        IOUtils.setByteArrayMaxOverride(200_000_000);
        try (InputStream is = file.getInputStream()) {
            // Sử dụng XSSFReader để đọc file Excel lớn theo kiểu streaming
            OPCPackage pkg = OPCPackage.open(is);
            XSSFReader reader = new XSSFReader(pkg);
            SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();
            XMLReader parser = XMLReaderFactory.createXMLReader();

            // Dùng ExecutorService để quản lý các luồng
            ExecutorService executor = Executors.newFixedThreadPool(THREAD_COUNT);

            // Phân chia công việc thành các phần nhỏ để xử lý song song
            List<Future<Void>> futures = new ArrayList<>();
            for (int i = 0; i < THREAD_COUNT; i++) {
                final int batchIndex = i;
                futures.add(executor.submit(() -> {
                    processSheet(reader, sst, batchIndex);
                    return null;
                }));
            }

            // Chờ tất cả các luồng hoàn thành
            for (Future<Void> future : futures) {
                future.get();
            }

            executor.shutdown();
        } catch (IOException | OpenXML4JException | SAXException | InterruptedException | ExecutionException e) {
            throw new RuntimeException("Error importing products from file", e);
        }
    }

    private void processSheet(XSSFReader reader, SharedStringsTable sst, int batchIndex) {
        try {
            // Đọc dữ liệu từ sheet theo từng phần
            InputStream sheet = reader.getSheetsData().next();
            XMLReader parser = XMLReaderFactory.createXMLReader();
            ProductSheetHandler handler = new ProductSheetHandler(sst, productRepository, batchIndex);
            parser.setContentHandler(handler);

            // Xử lý từng phần của sheet (chia ra theo batchIndex)
            parser.parse(new InputSource(sheet));
        } catch (Exception e) {
            throw new RuntimeException("Error processing sheet", e);
        }
    }

    private static class ProductSheetHandler extends DefaultHandler {
        private final SharedStringsTable sst;
        private final ProductRepository productRepository;
        private final int batchIndex;

        private String lastContents;
        private boolean nextIsString;
        private int col = 0;
        private String name;
        private double price;
        private int quantity;
        private List<Product> batch = new ArrayList<>();
        private static final int BATCH_SIZE = 1000;

        public ProductSheetHandler(SharedStringsTable sst, ProductRepository productRepository, int batchIndex) {
            this.sst = sst;
            this.productRepository = productRepository;
            this.batchIndex = batchIndex;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if ("c".equals(qName)) {
                // Xác định kiểu dữ liệu của cell
                String cellType = attributes.getValue("t");
                nextIsString = "s".equals(cellType);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            if ("v".equals(qName)) {
                // Lấy giá trị của cell
                String value = nextIsString ? sst.getItemAt(Integer.parseInt(lastContents)).getString() : lastContents;

                switch (col) {
                    case 1 -> name = value;
                    case 2 -> price = Double.parseDouble(value);
                    case 3 -> quantity = Integer.parseInt(value);
                }
                col++;
            } else if ("row".equals(qName)) {
                // Lưu dữ liệu vào batch
                if (col == 4) {
                    Product product = new Product(null, name, price, quantity);
                    batch.add(product);

                    // Nếu batch đủ lớn, lưu vào DB
                    if (batch.size() >= BATCH_SIZE) {
                        productRepository.saveAll(batch);
                        batch.clear();
                    }
                }
                col = 1; // reset cột
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
}
