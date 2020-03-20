import java.io.*;
import java.net.InetSocketAddress;
import java.util.Iterator;
import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;
import com.sun.net.httpserver.HttpServer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;


public class xlsx {
    private static final String FILE_NAME = "Test.xlsx";


    public static void main(String[] args) throws IOException {
        System.out.println("Server Started");
        HttpServer server = HttpServer.create(new InetSocketAddress(8000), 0);
        server.createContext("/xlsx", new LinkHandler());
        server.setExecutor(null);
        server.start();
    }

    static class LinkHandler implements HttpHandler {

        @Override
        public void handle(HttpExchange httpExchange) throws IOException {
            String response = "";
            httpExchange.sendResponseHeaders(200, response.length());
            InputStreamReader isr = new InputStreamReader(httpExchange.getRequestBody(), "utf-8");
            BufferedReader br = new BufferedReader(isr);
            String query = br.readLine();
            //System.out.println(query);
            min(query);


            response = x();
            OutputStream os = httpExchange.getResponseBody();
            os.write(response.getBytes());
            //System.out.println(response);
            os.close();
        }

        private static String x() throws IOException {
            String filePath = "Test.xlsx";
            File file = new File(filePath);
            FileInputStream fis = new FileInputStream(file);
            byte[] bytes = new byte[(int) file.length()];
            fis.read(bytes);
            String base64 = new sun.misc.BASE64Encoder().encode(bytes);


            fis.close();
            return base64;
        }

        private static void min(String query) throws IOException {
            //System.out.println(json);
            int rowNum = 0;
            int colNum = 0;

            JSONParser jsonParser = new JSONParser();
            String str = "Даты событий";

            try {
                JSONObject jsonObject = (JSONObject) jsonParser.parse(query);
                JSONArray jsonArray = (JSONArray) jsonObject.get("Сводка");
                JSONArray jsonArray2 = (JSONArray) jsonObject.get("События");

                Iterator<Object> iterator = jsonArray.iterator();
                Iterator<Object> iterator2 = jsonArray2.iterator();

                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet1 = workbook.createSheet("Сводка");
                XSSFSheet sheet2 = workbook.createSheet("События");
                Row row = sheet1.createRow(rowNum++);


                while (iterator.hasNext()) {
                    jsonObject = (JSONObject) iterator.next();

                    for (Object key : jsonObject.keySet()) {
                        if (key.equals(str)) // работа с датами событий
                        {
                            jsonArray = (JSONArray) jsonObject.get(key);

                            for(Object jb : jsonArray)
                            {
                                if(colNum == 0)
                                {
                                    Cell cell1 = row.createCell(colNum++);
                                    cell1.setCellValue("Фамилия");
                                    Cell cell2 = row.createCell(colNum++);
                                    cell2.setCellValue("Имя");
                                }
                                Cell cell = row.createCell(colNum++);
                                cell.setCellValue((String) jb);
                            }
                            colNum = 0;
                            row = sheet1.createRow(rowNum++);
                        }else  // работа с объектами-людьми
                        {
                            if (key.equals("Участие в событиях")) {
                                jsonArray = (JSONArray) jsonObject.get(key);

                                for (Object jb : jsonArray) {
                                    Cell cell = row.createCell(colNum++);
                                    cell.setCellValue((String) jb);
                                }
                                colNum = 0;
                                row = sheet1.createRow(rowNum++);
                            } else {
                                Cell cell = row.createCell(colNum++);
                                cell.setCellValue((String) jsonObject.get(key));

                            }
                        }
                    }

                }
                rowNum = 0;
                colNum = 0;
                Row row2 = sheet2.createRow(rowNum++);
                Cell cell;
                while (iterator2.hasNext()) {
                    jsonObject = (JSONObject) iterator2.next();
                    for (Object key : jsonObject.keySet()) {
                        //System.out.println(key + ":" + jsonObject.get(key));
                        cell = row2.createCell(colNum++);
                        cell.setCellValue((String) key);
                        jsonArray2 = (JSONArray) jsonObject.get(key);
                        for(Object jb : jsonArray2)
                        {
                            cell = row2.createCell(colNum++);
                            cell.setCellValue((String) jb);
                        }
                        colNum = 0;
                        row2 = sheet2.createRow(rowNum++);
                    }
                }
                try {
                    FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
                    workbook.write(outputStream);
                    workbook.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }

                System.out.println("Done");
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }
    }
}
