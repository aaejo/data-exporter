package io.github.aaejo.dataexporter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.Optional;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import io.micrometer.common.util.StringUtils;
import lombok.extern.slf4j.Slf4j;

/**
 * @author Jeffery Kung
 */
@Slf4j
@Service
public class DataExporter {

    @Value("${aaejo.jds.data-exporter.mark-deletions}")
    private boolean markDeletions;
    private final JdbcTemplate jdbcTemplate;

    public DataExporter(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public String retrieveData(Optional<String> password) throws Exception {
        String sql = "use jds;";
        String sqlForSameRows = "select distinct of.* from originalfile of left join scrapeddata sd on of.primeEmail = sd.primeEmail and of.personAttribute = sd.personAttribute where of.primeEmail = sd.primeEmail and of.address1 = sd.address1 and of.postalCode = sd.postalCode and of.institution = sd.institution and of.department = sd.department;";
        String sqlForChangeRows = "select distinct of.personID, of.salutation, of.fname, of.mname, of.lname, sd.address1, of.address2, of.address3, of.city, of.state, sd.postalCode, of.country, sd.department, of.institution, of.institutionId, of.primeEmail, of.userID, of.ORCID, of.ORCIDVal, of.personAttribute, of.memberStatus from originalfile of left join scrapeddata sd on of.primeEmail = sd.primeEmail and of.personAttribute = sd.personAttribute where of.primeEmail = sd.primeEmail and of.address1 not like sd.address1 or of.primeEmail = sd.primeEmail and of.postalCode not like sd.postalCode;";
        String sqlForDeleteRows = "select of.* from originalfile of where of.primeEmail not in (select sd.primeEmail from scrapeddata sd) or of.primeEmail in (select sd.primeEmail from scrapeddata sd) and of.personAttribute not in (select sd.personAttribute from scrapeddata sd left join originalfile of on sd.primeEmail = of.primeEmail where sd.primeEmail = of.primeEmail);";
        String sqlForNewRows = "select distinct sd.* from scrapeddata sd where sd.primeEmail not in (select of.primeEmail from originalfile of) or sd.primeEmail in (select of.primeEmail from originalfile of) and sd.personAttribute not in (select of.personAttribute from originalfile of where of.primeEmail = sd.primeEmail);";
        File dataFile = File.createTempFile("NewDIAUsersExport", ".xlsx");

        try (Connection connection = jdbcTemplate.getDataSource().getConnection();
                PreparedStatement ps = connection.prepareStatement(sql);
                PreparedStatement psSame = connection.prepareStatement(sqlForSameRows);
                PreparedStatement psChange = connection.prepareStatement(sqlForChangeRows);
                PreparedStatement psDelete = connection.prepareStatement(sqlForDeleteRows);
                PreparedStatement psNew = connection.prepareStatement(sqlForNewRows);

                ResultSet rs = ps.executeQuery();
                ResultSet rsSame = psSame.executeQuery();
                ResultSet rsChange = psChange.executeQuery();
                ResultSet rsDelete = psDelete.executeQuery();
                ResultSet rsNew = psNew.executeQuery()) {

            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("NewDIAUsersExport");

            int rownum = 0;

            Row row1 = sheet.createRow(rownum++);
            CellStyle style1 = workbook.createCellStyle();
            style1.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
            style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            int cellnum1 = 0;
            Cell cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Person ID");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Salutation");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("First Name");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Middle Name");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Last Name");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Address1");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Address2");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Address3");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("City");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("State/Province");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Postal Code");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Country/Region");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Department");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Institution");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Institution Identifier");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Primary E-mail Address");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("User ID");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("ORCID");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("ORCID Validation");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Person Attribute");
            cell1 = row1.createCell(cellnum1++);
            cell1.setCellStyle(style1);
            cell1.setCellValue("Member Status");

            while (rsSame.next()) {
                String personID, salutation, fname, mname, lname, address1, address2, address3, city, stateProv, postal,
                        countryRegion, department, institution, institutionID, primeEmail, userID, ORCID, ORCIDVal,
                        personAttribute, memberStatus;

                personID = rsSame.getString("personId");
                salutation = rsSame.getString("salutation");
                fname = rsSame.getString("fname");
                mname = rsSame.getString("mname");
                lname = rsSame.getString("lname");
                address1 = rsSame.getString("address1");
                address2 = rsSame.getString("address2");
                address3 = rsSame.getString("address3");
                city = rsSame.getString("city");
                stateProv = rsSame.getString("state");
                postal = rsSame.getString("postalCode");
                countryRegion = rsSame.getString("country");
                department = rsSame.getString("department");
                institution = rsSame.getString("institution");
                institutionID = rsSame.getString("institutionId");
                primeEmail = rsSame.getString("primeEmail");
                userID = rsSame.getString("userId");
                ORCID = rsSame.getString("ORCID");
                ORCIDVal = rsSame.getString("ORCIDVal");
                personAttribute = rsSame.getString("personAttribute");
                memberStatus = rsSame.getString("memberStatus");

                Row row = sheet.createRow(rownum++);
                int cellnum = 0;
                Cell cell = row.createCell(cellnum++);
                cell.setCellValue(personID);
                cell = row.createCell(cellnum++);
                cell.setCellValue(salutation);
                cell = row.createCell(cellnum++);
                cell.setCellValue(fname);
                cell = row.createCell(cellnum++);
                cell.setCellValue(mname);
                cell = row.createCell(cellnum++);
                cell.setCellValue(lname);
                cell = row.createCell(cellnum++);
                cell.setCellValue(address1);
                cell = row.createCell(cellnum++);
                cell.setCellValue(address2);
                cell = row.createCell(cellnum++);
                cell.setCellValue(address3);
                cell = row.createCell(cellnum++);
                cell.setCellValue(city);
                cell = row.createCell(cellnum++);
                cell.setCellValue(stateProv);
                cell = row.createCell(cellnum++);
                cell.setCellValue(postal);
                cell = row.createCell(cellnum++);
                cell.setCellValue(countryRegion);
                cell = row.createCell(cellnum++);
                cell.setCellValue(department);
                cell = row.createCell(cellnum++);
                cell.setCellValue(institution);
                cell = row.createCell(cellnum++);
                cell.setCellValue(institutionID);
                cell = row.createCell(cellnum++);
                cell.setCellValue(primeEmail);
                cell = row.createCell(cellnum++);
                cell.setCellValue(userID);
                cell = row.createCell(cellnum++);
                cell.setCellValue(ORCID);
                cell = row.createCell(cellnum++);
                cell.setCellValue(ORCIDVal);
                cell = row.createCell(cellnum++);
                cell.setCellValue(personAttribute);
                cell = row.createCell(cellnum++);
                cell.setCellValue(memberStatus);

            }

            while (rsChange.next()) {
                String personID, salutation, fname, mname, lname, address1, address2, address3, city, stateProv, postal,
                        countryRegion, department, institution, institutionID, primeEmail, userID, ORCID, ORCIDVal,
                        personAttribute, memberStatus;

                personID = rsChange.getString("personId");
                salutation = rsChange.getString("salutation");
                fname = rsChange.getString("fname");
                mname = rsChange.getString("mname");
                lname = rsChange.getString("lname");
                address1 = rsChange.getString("address1");
                address2 = rsChange.getString("address2");
                address3 = rsChange.getString("address3");
                city = rsChange.getString("city");
                stateProv = rsChange.getString("state");
                postal = rsChange.getString("postalCode");
                countryRegion = rsChange.getString("country");
                department = rsChange.getString("department");
                institution = rsChange.getString("institution");
                institutionID = rsChange.getString("institutionId");
                primeEmail = rsChange.getString("primeEmail");
                userID = rsChange.getString("userId");
                ORCID = rsChange.getString("ORCID");
                ORCIDVal = rsChange.getString("ORCIDVal");
                personAttribute = rsChange.getString("personAttribute");
                memberStatus = rsChange.getString("memberStatus");

                Row row = sheet.createRow(rownum++);
                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setColor(IndexedColors.BLUE.getIndex());
                style.setFont(font);
                int cellnum = 0;
                Cell cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(personID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(salutation);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(fname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(mname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(lname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address1);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address2);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address3);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(city);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(stateProv);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(postal);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(countryRegion);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(department);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(institution);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(institutionID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(primeEmail);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(userID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(ORCID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(ORCIDVal);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(personAttribute);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(memberStatus);

            }

            while (rsNew.next()) {
                String personID, salutation, fname, mname, lname, address1, address2, address3, city, stateProv, postal,
                        countryRegion, department, institution, institutionID, primeEmail, userID, ORCID, ORCIDVal,
                        personAttribute, memberStatus;

                personID = "";
                salutation = rsNew.getString("salutation");
                fname = rsNew.getString("fname");
                mname = rsNew.getString("mname");
                lname = rsNew.getString("lname");
                address1 = rsNew.getString("address1");
                address2 = rsNew.getString("address2");
                address3 = rsNew.getString("address3");
                city = rsNew.getString("city");
                stateProv = rsNew.getString("state");
                postal = rsNew.getString("postalCode");
                countryRegion = rsNew.getString("country");
                department = rsNew.getString("department");
                institution = rsNew.getString("institution");
                institutionID = "";
                primeEmail = rsNew.getString("primeEmail");
                userID = rsNew.getString("userId");
                ORCID = "";
                ORCIDVal = "";
                personAttribute = rsNew.getString("personAttribute");
                memberStatus = "";

                Row row = sheet.createRow(rownum++);
                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setColor(IndexedColors.GREEN.getIndex());
                style.setFont(font);
                int cellnum = 0;
                Cell cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(personID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(salutation);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(fname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(mname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(lname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address1);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address2);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address3);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(city);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(stateProv);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(postal);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(countryRegion);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(department);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(institution);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(institutionID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(primeEmail);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(userID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(ORCID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(ORCIDVal);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(personAttribute);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(memberStatus);

            }

            while (markDeletions && rsDelete.next()) {
                String personID, salutation, fname, mname, lname, address1, address2, address3, city, stateProv, postal,
                        countryRegion, department, institution, institutionID, primeEmail, userID, ORCID, ORCIDVal,
                        personAttribute, memberStatus;

                personID = rsDelete.getString("personId");
                salutation = rsDelete.getString("salutation");
                fname = rsDelete.getString("fname");
                mname = rsDelete.getString("mname");
                lname = rsDelete.getString("lname");
                address1 = rsDelete.getString("address1");
                address2 = rsDelete.getString("address2");
                address3 = rsDelete.getString("address3");
                city = rsDelete.getString("city");
                stateProv = rsDelete.getString("state");
                postal = rsDelete.getString("postalCode");
                countryRegion = rsDelete.getString("country");
                department = rsDelete.getString("department");
                institution = rsDelete.getString("institution");
                institutionID = rsDelete.getString("institutionId");
                primeEmail = rsDelete.getString("primeEmail");
                userID = rsDelete.getString("userId");
                ORCID = rsDelete.getString("ORCID");
                ORCIDVal = rsDelete.getString("ORCIDVal");
                personAttribute = rsDelete.getString("personAttribute");
                memberStatus = rsDelete.getString("memberStatus");

                Row row = sheet.createRow(rownum++);
                CellStyle style = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setColor(IndexedColors.RED.getIndex());
                style.setFont(font);
                int cellnum = 0;
                Cell cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(personID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(salutation);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(fname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(mname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(lname);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address1);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address2);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(address3);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(city);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(stateProv);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(postal);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(countryRegion);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(department);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(institution);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(institutionID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(primeEmail);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(userID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(ORCID);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(ORCIDVal);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(personAttribute);
                cell = row.createCell(cellnum++);
                cell.setCellStyle(style);
                cell.setCellValue(memberStatus);

            }

            for (int i = 0; i < 21; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream out = new FileOutputStream(dataFile)) {
                workbook.write(out);
            }
            workbook.close();

            log.info("Excel written successfully.");

            if (password.isEmpty() || StringUtils.isBlank(password.get())) {
                log.info("No password provided, skipping encryption.");
                return dataFile.getAbsolutePath();
            }

            try (POIFSFileSystem fs = new POIFSFileSystem()) {
                EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
                Encryptor enc = info.getEncryptor();
                enc.confirmPassword(password.get());
                ZipSecureFile.setMinInflateRatio(0);
                try (OPCPackage opc = OPCPackage.open(dataFile, PackageAccess.READ_WRITE);
                        OutputStream os = enc.getDataStream(fs)) {
                    opc.save(os);
                    opc.close();
                }
                ZipSecureFile.setMinInflateRatio(0);
                try (FileOutputStream fos = new FileOutputStream(dataFile)) {
                    fs.writeFilesystem(fos);
                    fs.close();
                }
            }

            log.info("Excel encrypted successfully.");
            return dataFile.getAbsolutePath();
        }
    }
}
