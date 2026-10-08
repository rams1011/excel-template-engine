package com.tx.excel.template;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.net.URL;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertNotNull;

class ExcelGenTest {

    @TempDir
    Path tempDir;

    @Test
    void testExcelGeneration() throws Exception {
        Map<String, Object> data = Map.of(
                "title", "Payroll",
                "company", Map.of("name", "Acme"),
                "period", "Sep 2026",
                "printedBy", "Ram",
                "total", 1250.5,
                "items", List.of(
                        Map.of("name", "Ada", "amount", 700),
                        Map.of("name", "Bob", "amount", 550.5)
                )
        );

        ExcelTemplateEngine.render(
                Path.of("src/test/resources/templates/sample-report.xlsx"),
                Path.of("src/test/resources/templates/sample-output.xlsx"),
                data);
    }

    @Test
    void testUATDataGeneration() throws Exception {
        Map<String, Object> data = Map.of(
                "reportLabel", "QuickSolver-Report#",
                "title", "Massachusetts Department of Industrial Accidents – Peril 13",
                "runDateLine", "Run Date: 06/26/2026 04:04:55",
                "periodLine", "Quarter 4 2025; 10/01/2025 to 12/31/2025",
                "totalAssessment", 9706.0,
                "rows", List.of(
                        peril13PolicyRow(
                                "0021835-07-774103",
                                "WISCONSIN EVANGELICAL LUTHERAN SYNOD",
                                "N16W23377 STONE RIDGE DR  WAUKESHA WI 53188-1108",
                                "064",
                                -30.42,
                                -1.0,
                                "Premium Audit"
                        ),
                        peril13PolicyRow(
                                "0021835-07-774103",
                                "WISCONSIN EVANGELICAL LUTHERAN SYNOD",
                                "N16W23377 STONE RIDGE DR  WAUKESHA WI 53188-1108",
                                "064",
                                400.14,
                                19.0,
                                "Premium Audit"
                        )
                )
        );

        ExcelTemplateEngine.render(
                Path.of("D:\\code\\excel-template-engine\\samples\\ma-dia-peril-13-template.xlsx"),
                Path.of("D:\\code\\excel-template-engine\\samples\\ma-dia-peril-13-output.xlsx"),
                data);
    }

    private static Map<String, Object> peril13PolicyRow(
            String policyNumber,
            String policyHolderName,
            String policyHolderAddress,
            String location,
            double standardPremium,
            double assessment,
            String assessmentReason) {
        return Map.ofEntries(
                Map.entry("companyCode", "CMIC"),
                Map.entry("naic", "18767"),
                Map.entry("policyNumber", policyNumber),
                Map.entry("policyHolderName", policyHolderName),
                Map.entry("policyHolderAddress", policyHolderAddress),
                Map.entry("fein", "00-0000000"),
                Map.entry("policyEffectiveDate", 45474),
                Map.entry("location", location),
                Map.entry("standardPremium", standardPremium),
                Map.entry("effectiveDiaRate", 0.0468),
                Map.entry("assessment", assessment),
                Map.entry("assessmentReason", assessmentReason)
        );
    }

    private static Path classpathResource(String resourcePath) throws Exception {
        URL resource = ExcelGenTest.class.getResource(resourcePath);
        assertNotNull(resource, resourcePath + " must be on the test classpath");
        return Path.of(resource.toURI());
    }
}
