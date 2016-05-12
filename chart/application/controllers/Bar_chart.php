<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Bar_chart extends CI_Controller {


	public function index()
	{


		$this->load->library('excel');
		$objPHPExcel = new PHPExcel();
		$objWorksheet = $objPHPExcel->getActiveSheet();
		$objWorksheet->fromArray(
			array(
				array('',       'Rainfall (mm)',    'Temperature (°F)', 'Humidity (%)'),
				array('Jan',        78,                 52,                 61),
				array('Feb',        64,                 54,                 62),
				array('Mar',        62,                 57,                 63),
				array('Apr',        21,                 62,                 59),
				array('May',        11,                 75,                 60),
				array('Jun',        1,                  75,                 57),
				array('Jul',        1,                  79,                 56),
				array('Aug',        1,                  79,                 59),
				array('Sep',        10,                 75,                 60),
				array('Oct',        40,                 68,                 63),
				array('Nov',        69,                 62,                 64),
				array('Dec',        89,                 57,                 66),
				)
			);


//  Set the Labels for each data series we want to plot
//      Datatype
//      Cell reference for data
//      Format Code
//      Number of datapoints in series
//      Data values
//      Data Marker
		$dataseriesLabels1 = array(
    new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$1', NULL, 1),   //  Temperature
    );
		$dataseriesLabels2 = array(
    new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$C$1', NULL, 1),   //  Rainfall
    );
		$dataseriesLabels3 = array(
    new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$D$1', NULL, 1),   //  Humidity
    );

//  Set the X-Axis Labels
//      Datatype
//      Cell reference for data
//      Format Code
//      Number of datapoints in series
//      Data values
//      Data Marker

		$xAxisTickValues = array(
    new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2:$A$13', NULL, 12),    //  Jan to Dec
    );


//  Set the Data values for each data series we want to plot
//      Datatype
//      Cell reference for data
//      Format Code
//      Number of datapoints in series
//      Data values
//      Data Marker

		$dataSeriesValues1 = array(
			new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$2:$B$13', NULL, 12),
			);

//  Build the dataseries
		$series1 = new PHPExcel_Chart_DataSeries(
    PHPExcel_Chart_DataSeries::TYPE_BARCHART,       // plotType
    PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,  // plotGrouping
    range(0, count($dataSeriesValues1)-1),          // plotOrder
    $dataseriesLabels1,                             // plotLabel
    $xAxisTickValues,                               // plotCategory
    $dataSeriesValues1                              // plotValues
    );
//  Set additional dataseries parameters
//      Make it a vertical column rather than a horizontal bar graph
		$series1->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);


//  Set the Data values for each data series we want to plot
//      Datatype
//      Cell reference for data
//      Format Code
//      Number of datapoints in series
//      Data values
//      Data Marker
		$dataSeriesValues2 = array(
			new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$2:$C$13', NULL, 12),
			);

//  Build the dataseries
		$series2 = new PHPExcel_Chart_DataSeries(
    PHPExcel_Chart_DataSeries::TYPE_LINECHART,      // plotType
    PHPExcel_Chart_DataSeries::GROUPING_STANDARD,   // plotGrouping
    range(0, count($dataSeriesValues2)-1),          // plotOrder
    $dataseriesLabels2,                             // plotLabel
    NULL,                                           // plotCategory
    $dataSeriesValues2                              // plotValues
    );


//  Set the Data values for each data series we want to plot
//      Datatype
//      Cell reference for data
//      Format Code
//      Number of datapoints in series
//      Data values
//      Data Marker
		$dataSeriesValues3 = array(
			new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$D$2:$D$13', NULL, 12),
			);

//  Build the dataseries
		$series3 = new PHPExcel_Chart_DataSeries(
    PHPExcel_Chart_DataSeries::TYPE_AREACHART,      // plotType
    PHPExcel_Chart_DataSeries::GROUPING_STANDARD,   // plotGrouping
    range(0, count($dataSeriesValues2)-1),          // plotOrder
    $dataseriesLabels3,                             // plotLabel
    NULL,                                           // plotCategory
    $dataSeriesValues3                              // plotValues
    );


//  Set the series in the plot area
		$plotarea = new PHPExcel_Chart_PlotArea(NULL, array($series1, $series2, $series3));
//  Set the chart legend
		$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);

		$title = new PHPExcel_Chart_Title('Average Weather Chart');


//  Create the chart
		$chart = new PHPExcel_Chart(
    'chart1',       // name
    $title,         // title
    $legend,        // legend
    $plotarea,      // plotArea
    true,           // plotVisibleOnly
    0,              // displayBlanksAs
    NULL,           // xAxisLabel
    NULL            // yAxisLabel
    );

//  Set the position where the chart should appear in the worksheet
		$chart->setTopLeftPosition('F2');
		$chart->setBottomRightPosition('O16');

//  Add the chart to the worksheet
		$objWorksheet->addChart($chart);

// Save Excel 2007 file
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="sample.xlsx"');
		header('Cache-Control: max-age=0');

		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->setIncludeCharts(TRUE);
		$objWriter->save('php://output');
	}
}
