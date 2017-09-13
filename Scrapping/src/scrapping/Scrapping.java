package scrapping;

import java.io.File;
import java.io.FileOutputStream;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.IOException;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileSystemView;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class Scrapping {

    public static void main(String[] args) {
        JOptionPane.showMessageDialog(null, "Pulse Aceptar para crear Excel. (Durara alrededor de 2 minutos)");
        //ARCHIVO EXCEL
        try{
        FileOutputStream archivo = null;

        File home = FileSystemView.getFileSystemView().getHomeDirectory();//El fichero EXCEL se creara en el escritorio
        String ruta = home.getAbsolutePath() + "/PreciosWebs.xls";
        File archivoXLS = new File(ruta);
        if (archivoXLS.exists()) {
            archivoXLS.delete(); //SI EXISTIA LA BORRO
        }
        archivoXLS.createNewFile(); // CREO UN NUEVO ARCHIVO XLS
        Workbook libro = new HSSFWorkbook();
        archivo = new FileOutputStream(archivoXLS);
        org.apache.poi.ss.usermodel.Sheet hoja = libro.createSheet("Precios Webs");//Creamos una hoja en el EXCEL
		
        HSSFCellStyle estiloCelda = (HSSFCellStyle) libro.createCellStyle();//definimos un estilo de celda
        estiloCelda.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        estiloCelda.setFillForegroundColor((short) 22);
        estiloCelda.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		
        HSSFCellStyle estiloCelda2 = (HSSFCellStyle) libro.createCellStyle();//definimos otro estilo de celda
        estiloCelda2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		
        Row fila = hoja.createRow(0);//Creamos una fila, que sera la cabecera del EXCEL
        for (int c = 0; c < 3; c++) { //CABECERA DE LA HOJA EXCEL
            org.apache.poi.ss.usermodel.Cell celda = fila.createCell(c);//Vamos creando de uno en uno cada celda de la primera fila (Cabecera)
            celda.setCellStyle(estiloCelda);
            if (c == 0) {
                celda.setCellValue("TIENDA");
            }
            if (c == 1) {
                celda.setCellValue("PRODUCTO");
            }
            if (c == 2) {
                celda.setCellValue("PRECIO");
            }
        }
        int x = 1;//Variable de control de filas para el EXCEL
        //VOY OBTENIENDO LOS DATOS Y METIENDOLOS EN EL EXCEL

        Document doc;
        int maxBodySize = 4092000;//Para paginas webs muy grandes, este valor hay que aumentarlo para que cargue toda la pagina completa
        
        
        //Carrefour
        doc = Jsoup.connect("https://www.carrefour.es/global/?Dy=1&Nty=1&Ntx=mode+matchallany&Ntt=usisa&search=").maxBodySize(maxBodySize).get();
        String title = "CARREFOUR";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);//Creo una nueva fila para poner el nombre de la pagina web
        org.apache.poi.ss.usermodel.Cell celda = fila.createCell(0);//lo inserto en la primera celda
        celda.setCellValue(title);
        x++;
        Elements linksNombre = doc.select("h2.titular-producto");//Obtengo las etiquetas html: h2 cuya clase sea 'titular-producto'(Hay estan los nombres de los productos)
        Elements linksPrecios = doc.select("p.precio-nuevo");//Obtengo las etiquetas html: p cuya clase sea 'precio-nuevo'
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(i).text());
            fila = hoja.createRow(x);//Voy creando filas a medida que obtengo los datos
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);// celda para poner el nombre
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);//celda para poner el precio
            celda2.setCellValue(linksPrecios.get(i).text());
            x++;
        }
		/* EL RESTO DEL CODIGO SE REPITE, PERO PARA DISTINTAS PAGINAS WEBS
        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("https://www.carrefour.es/global/?Dy=1&Nty=1&Ntx=mode+matchallany&Ntt=conservas+tejero&search=").maxBodySize(maxBodySize).get();
        title = "CARREFOUR";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        linksNombre = doc.select("h2.titular-producto");
        linksPrecios = doc.select("p.precio-nuevo");
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(i).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(i).text());
            x++;
        }

        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("http://www.hechoenandalucia.net/buscar?controller=search&orderby=position&orderway=desc&search_query=tejero").maxBodySize(maxBodySize).get();
        title = "HECHOENANDALUCIA";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        linksNombre = doc.select("p.product-desc");
        linksPrecios = doc.select("span.product-price");
        int j = 0;
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(j).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(j).text());
            x++;
            j = j + 2;

        }

        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("https://www.alcampo.es/compra-online/search/?text=tejero").maxBodySize(maxBodySize).get();
        title = "ALCAMPO";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        linksNombre = doc.select("div.productName");
        linksPrecios = doc.select("span.price");
        j = 1;
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(j).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(j).text());
            x++;
            j++;
        }

        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("https://www.alcampo.es/compra-online/search/?text=usisa").maxBodySize(maxBodySize).get();
        title = "ALCAMPO";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        linksNombre = doc.select("div.productName");
        linksPrecios = doc.select("span.price");
        j = 1;
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(j).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(j).text());
            x++;
            j++;
        }

        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("http://ultramarinosalonso.com/es/buscar?controller=search&orderby=position&orderway=desc&search_query=usisa&submit_search=").maxBodySize(maxBodySize).get();
        title = "ULTRAMARINOSALONSO";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        linksNombre = doc.select("p.product-desc");
        linksPrecios = doc.select("span.product-price");
        j = 0;
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(j).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(j).text());
            x++;
            j = j + 2;
        }

        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("http://ultramarinosalonso.com/es/buscar?controller=search&orderby=position&orderway=desc&search_query=tejero&submit_search=").maxBodySize(maxBodySize).get();
        title = "ULTRAMARINOSALONSO";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        linksNombre = doc.select("p.product-desc");
        linksPrecios = doc.select("span.product-price");
        j = 0;
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(j).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(j).text());
            x++;
            j = j + 2;
        }

        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("https://www.salmantina.com/catalogsearch/result/?q=usisa").maxBodySize(maxBodySize).get();
        title = "SALMANTINA";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        linksNombre = doc.select("h2.product-name");
        linksPrecios = doc.select("span.price");
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(i).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(i).text());
            x++;
        }

        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("https://www.salmantina.com/catalogsearch/result/?q=tejero").maxBodySize(maxBodySize).get();
        title = "SALMANTINA";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        linksNombre = doc.select("h2.product-name");
        linksPrecios = doc.select("span.price");
        for (int i = 0; i < linksNombre.size(); i++) {

            // get the value from href attribute
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(i).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(i).text());
            x++;
        }

        System.out.println();
        System.out.println();
        System.out.println("------------------");
        System.out.println();
        System.out.println();
        doc = Jsoup.connect("https://www.amazon.es/s/ref=nb_sb_noss_1?__mk_es_ES=%C3%85M%C3%85%C5%BD%C3%95%C3%91&url=search-alias%3Daps&field-keywords=conservas+tejero").maxBodySize(maxBodySize).get();
        title = "AMAZON";
        System.out.println("title : " + title);
        fila = hoja.createRow(x);
        celda = fila.createCell(0);
        celda.setCellValue(title);
        x++;
        Elements result = doc.select("h2#s-result-count");
        linksNombre = doc.select("h2.s-access-title");
        linksPrecios = doc.select("span.s-price");
        int fin = linksNombre.size();
        if(!result.get(0).text().substring(0, 2).replace(" ", "").equals("Re")){
            fin = Integer.parseInt(result.get(0).text().substring(0, 2).replace(" ", ""));
        }

        for (int i = 0; i < fin; i++) {
            System.out.println("Nombre : " + linksNombre.get(i).text());
            System.out.println("Precio : " + linksPrecios.get(i).text());
            fila = hoja.createRow(x);
            org.apache.poi.ss.usermodel.Cell celda1 = fila.createCell(1);
            celda1.setCellValue(linksNombre.get(i).text());
            org.apache.poi.ss.usermodel.Cell celda2 = fila.createCell(2);
            celda2.setCellValue(linksPrecios.get(i).text());
            x++;
        }
        
        */
        libro.write(archivo);
        }catch(Exception e){
            JOptionPane.showMessageDialog(null, "Error al crear Excel.\nError: "+e.getMessage());
        }
        JOptionPane.showMessageDialog(null, "Archivo Excel creado.");
        
    }


}
