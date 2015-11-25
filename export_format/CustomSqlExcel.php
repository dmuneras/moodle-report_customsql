<?php

// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.


/**
 *
 *
 * @package report_customsql
 * @copyright 2014 Daniel Munera based on 2009 The Open University
 * @license http://www.gnu.org/copyleft/gpl.html GNU GPL v3 or later
 */

/* EXPORT TO EXCEL
-----------------------------------------------------------------------------*/

require_once($CFG->dirroot .'/lib/excellib.class.php');
require_once($CFG->dirroot .'/lib/moodlelib.php');
require_once($CFG->dirroot .'/report/customsql/locallib.php');
require_once($CFG->dirroot .'/report/customsql/export_format/CustomSqlExportFormat.php');

/**
* CustomSqlExcel
*
*
* @package    report_customsql
* @subpackage export_format
* @author     Daniel Múnera Sánchez <dmunera119@gmail.com>
*/
class CustomSqlExcel implements CustomSqlExportFormat{

	/**
 	 * Function to send a SQL report in excel format
	 * @param $id identification of the custom SQL.
	 * @param  $csvtimestamp TIMESTAMP
	 * @param $params Params of the Custom SQL report. it is optional.
	 * @return stdClass Object SQL statement result
	 */
	public function send_report_customsql($id, $csvtimestamp, $params =array()){
		global $DB;
		$report = $DB->get_record('report_customsql_queries', array('id' => $id));
		if (!$report){
			print_error('invalidreportid', 'report_customsql',
	    		report_customsql_ptt_url('index.php'), $id);
		}
		$sql_result =
			report_customsql_get_report_result($report,$csvtimestamp,$params);

		$matrix_report =
			report_customsql_create_matrix($sql_result);

		$this->send_report_predefined($matrix_report, $csvtimestamp);
	}

	/**
 	 * Function to send a Predefined report in excel format
	 * @param $report Array with the result of the predefined report
	 *		  execution
	 * @param  $timenow TIMESTAMP
	 * @return stdClass Object SQL statement result
	 */
	public function send_report_predefined($report, $timenow,$coursename = ""){
		global $CFG;

	    $filename = 'report_'.(time()).'.xls';
	    $downloadfilename = clean_filename($filename);

	    /// Creating a workbook
	    $workbook = new MoodleExcelWorkbook("-");

	    //Format definition based on PLACE TO TRAIN style guidelines
	    $highGradeFormat = $workbook->add_format(
	    	array(
	    		"color" => "black",
	    		"bg_color" => "#C7EECF",
	    		"top" => 1,
	    		"bottom" => 1,
	    		"left" => 1,
	    		"right" => 1,
	    		"align" => 'center',
	    		"v_align" => 'center'

	    	));

	    $lowGradeFormat = $workbook->add_format(
	    	array(
	    		"bg_color" => "#FEC7CD",
	    		"top" => 1,
	    		"bottom" => 1,
	    		"left" => 1,
	    		"right" => 1,
	    		"align" => 'center',
	    		"v_align" => 'center'
	     		));

	    $averageGradeFormat = $workbook->add_format(
	    	array(
	    		"color" => "black",
	    		"bg_color" => "#FEEB9F",
				"top" => 1,
	    		"bottom" => 1,
	    		"left" => 1,
	    		"right" => 1,
	    		"align" => 'center',
	    		"v_align" => 'center'
	    		));
	    $superTitleFormat = $workbook->add_format(
	    	array(
	    		"bg_color" => "#0071B6",
	    		"color" => "white",
	    		"top" => 1,
	    		"bottom" => 1,
	    		"left" => 1,
	    		"right" => 1,
	    		"align" => "center",
	    		"bold" => 1,
	    		"size" => 26,
	    		"text_wrap" => 1,
	    		"v_align" => 'center'
	    	));

	    $titleFormat = $workbook->add_format(
	    	array(
	    		"color" => "#0071B6",
	    		"top" => 1,
	    		"bottom" => 1,
	    		"left" => 1,
	    		"right" => 1,
	    		"align" => "center",
	    		"bold" => 1,
	    		"size" => 12,
	    		"text_wrap" => 1,
	    		"v_align" => 'center'
	    	));

	    $normalFormat = $workbook->add_format(
	    	array(
	    		"top" => 1,
	    		"bottom" => 1,
	    		"left" => 1,
	    		"right" => 1,
	    		"align" => "equal_space",
	    		"text_wrap" => 1
	    	));


	    /// Sending HTTP headers
	    $workbook->send($downloadfilename);

	    /// Adding the worksheet
	    $myxls = $workbook->add_worksheet($coursename);
	    $myxls->write_blank(0,0);

	    //Fill worksheet with the report values

	    //variables to count the columns length
	    $number_of_columns = 0;


	    //$ri => Row index
	    //$col => Columns values
	    //$ci => Columns index
	    //$cv => Cell value
	    foreach($report as $ri=>$col){
	    	$delta_row = 2;
	        foreach($col as $ci=>$cv){
	        	$col = $ci+1;
	        	$row = $ri+ $delta_row;
	            if($ri == 0){
	            	$number_of_columns +=1;
	                $myxls->write_string($row,$col,$cv, $titleFormat);
	                continue;
	            }
	            //Chech if cell value is numeric to apply specific format
	            if(is_numeric($cv)){
	                $number = (float) $cv;
	                if($cv >= 80){
	                    $myxls->write_number($row,$col,$number, $highGradeFormat);
	                }elseif ($cv >= 70) {
	                    $myxls->write_number($row,$col,$number, $averageGradeFormat);
	                }else{
	                    $myxls->write_number($row,$col,$number, $lowGradeFormat);
	                }
	            }else{
	              $myxls->write_string($row,$col,$cv, $normalFormat);
	            }
	        }
	    }
	    //Add title on top of the worksheet
	    $report_title =  "Informe Semanal ". $coursename .
	    	". Fecha: " . userdate($timenow,"%d / %m / %Y");
	    $myxls->write_string(1,1,$report_title,$superTitleFormat);
	    $myxls->set_row(1,80);
	    $myxls->merge_cells(1,1,1,$number_of_columns);

	    //Ad Place to Train logo
	    $image = $CFG->dirroot . "/report/customsql/images/report_logo.png";
	    $myxls->insert_bitmap(2, 1, $image,5,5);

	    //Increase height of titles row
	    $myxls->set_row(2,70);

	    //Increase width of all columns
	    $myxls->set_column(1,$number_of_columns,30);
	    $workbook->close();
	}

}
