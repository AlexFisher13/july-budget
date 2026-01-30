package com.fisher.july_budget;

import com.fisher.july_budget.excel.ExcelAggregationService;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.InputStream;

@SpringBootTest
@AutoConfigureMockMvc
class ExcelAggregationServiceTest {

    private final ExcelAggregationService service = new ExcelAggregationService();

    @Test
    void aggregateTest() throws Exception {
        ClassPathResource resource = new ClassPathResource("excel/input.xlsx");

        try (InputStream is = resource.getInputStream()) {
            service.aggregateByCategory(is);
        }
    }
}
