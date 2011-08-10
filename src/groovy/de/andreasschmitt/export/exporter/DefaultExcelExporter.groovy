package de.andreasschmitt.export.exporter

import de.andreasschmitt.export.builder.ExcelBuilder

/**
 * @author Andreas Schmitt
 *
 */
class DefaultExcelExporter extends AbstractExporter {

    protected void exportData(OutputStream outputStream, List data, List fields) throws ExportingException{
        try {
            def builder = new ExcelBuilder()

            builder {
                workbook(outputStream: outputStream){
                    sheet(name:"Simule-RH" ?: "Export", widths: getParameters().get("column.widths")){
                        //Default format
                        format(name: "header"){
                            font(name: "arial", bold: true)
                        }
                        format(name: "orgaoNomeFont"){
                            font(name: "arial", bold: true)
                        }
                        format(name: "orgaoUgFont"){
                            font(name: "arial", bold: false)
                        }
                        format(name: "titleFont"){
                            font(name: "arial", bold: true)
                        }
                        format(name: "trabalhadoresFont"){
                            font(name: "arial", bold: false)
                        }
                        //Create title
                        cell(row: 0,column:1, value: getParameters().get("orgaoNome"), format: "orgaoNomeFont")
                        cell(row: 1,column:1, value: getParameters().get("orgaoUg"), format: "orgaoUgFont")
                        cell(row: 2,column:1, value: getParameters().get("title"), format: "titleFont")
                        cell(row: 3,column:1, value: getParameters().get("trabalhadores"), format: "trabalhadoresFont")
                        //Create header
                        fields.eachWithIndex { field, index ->
                            String value = getLabel(field)
                            cell(row: 5, column: index, value: value, format: "header")
                        }

                        //Rows
                        data.eachWithIndex { object, k ->
                            fields.eachWithIndex { field, i ->
                                Object value = getValue(object, field)
                                cell(row: k + 6, column: i, value: value)
                            }
                        }
                        cell(row: data.size()+6, column: 0, value: getParameters().get("footer"))
                    }
                }
            }

            builder.write()
        }
        catch(Exception e){
            throw new ExportingException("Error during export", e)
        }
    }

}