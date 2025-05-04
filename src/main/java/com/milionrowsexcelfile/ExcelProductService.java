package com.milionrowsexcelfile;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;

@Service
public class ExcelProductService {

    private final ProductRepository productRepository;

    public ExcelProductService(ProductRepository productRepository) {
        this.productRepository = productRepository;
    }

    public void importProducts(MultipartFile file) {
        InputStream inputStream = null;
        try {
            inputStream = file.getInputStream();
        OPCPackage pkg = null;

        pkg = OPCPackage.open(inputStream);
        XSSFReader reader = new XSSFReader(pkg);
        SharedStringsTable sst = (SharedStringsTable) reader.getSharedStringsTable();

        XMLReader parser = XMLReaderFactory.createXMLReader();
        ProductSheetHandler handler = new ProductSheetHandler(sst, productRepository);
        parser.setContentHandler(handler);

        InputStream sheet = reader.getSheetsData().next();
        parser.parse(new InputSource(sheet));
        } catch (IOException | OpenXML4JException | SAXException e) {
            throw new RuntimeException(e);
        }
    }
}
