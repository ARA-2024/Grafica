import scala.io.Source
import scala.jdk.CollectionConverters._
import org.knowm.xchart.{CategoryChart, CategoryChartBuilder, BitmapEncoder}
import org.knowm.xchart.BitmapEncoder.BitmapFormat
import java.io.{FileInputStream, FileOutputStream}
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.util.Units
import org.apache.poi.ss.usermodel.Workbook

object Main extends App {
  // Lee los datos del CSV
  val filename = "sueldos.csv"
  val lines = Source.fromFile(filename).getLines().drop(1).toList

  val ciudades = lines.map(_.split(",")(0))
  val sueldos = lines.map(_.split(",")(1).toInt)

  // Crea la gráfica
  val chart: CategoryChart = new CategoryChartBuilder()
    .width(1000).height(600)
    .title("Sueldo por Ciudad")
    .xAxisTitle("Ciudad")
    .yAxisTitle("Sueldo")
    .build()

  chart.addSeries("Sueldo", ciudades.asJava, sueldos.map(_.asInstanceOf[Number]).asJava)

  // Exporta la gráfica como imagen PNG
  val imgFile = "grafica_sueldos.png"
  BitmapEncoder.saveBitmap(chart, "grafica_sueldos", BitmapFormat.PNG)

  // Crea un documento DOCX e inserta la imagen
  val doc = new XWPFDocument()
  val out = new FileOutputStream("grafica_sueldos.docx")
  val paragraph = doc.createParagraph()
  val run = paragraph.createRun()

  val is = new FileInputStream(imgFile)
  run.addPicture(is, Workbook.PICTURE_TYPE_PNG, imgFile, Units.toEMU(500), Units.toEMU(300))
  is.close()

  doc.write(out)
  out.close()
  doc.close()

  println("¡Documento DOCX creado con la gráfica! Puedes abrirlo con WordPad o Word.")
}