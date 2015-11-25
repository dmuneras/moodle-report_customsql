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

/* EXPORT TO CSV
-----------------------------------------------------------------------------*/

require_once($CFG->dirroot.'/lib/excellib.class.php');
require_once($CFG->dirroot .'/report/customsql/locallib.php');
require_once($CFG->dirroot .
    '/report/customsql/export_format/CustomSqlExportFormat.php');

/**
* CustomSqlPttCSV
*
*
* @package    report_customsql_ptt
* @subpackage export_format
* @author     Daniel Múnera Sánchez <dmunera119@gmail.com>
*/
class CustomSqlCSV implements CustomSqlExportFormat{


    /**
     * Function to send a SQL report in CSV format
     * @param $id identification of the custom SQL.
     * @param  $csvtimestamp TIMESTAMP
     * @param $params Params of the Custom SQL report. it is optional.
     * @return stdClass Object SQL statement result
     */
	public function send_report_customsql($id , $csvtimestamp, $params =array()){
		global $DB;
		$report = $DB->get_record(
            'report_customsql_ptt_queries',
            array('id' => $id)
        );

		if (!$report) {
		    print_error('invalidreportid',
		    	'report_customsql_ptt', report_customsql_ptt_url('index.php'), $id);
		}

		list($csvfilename) = $this->csv_filename($report, $csvtimestamp);
		if (!is_readable($csvfilename)) {
		    print_error('unknowndownloadfile', 'report_customsql_ptt',
		                report_customsql_ptt_url('view.php?id=' . $id));
		}
		send_file(
            $csvfilename,
            'report.csv',
            'default' ,
            0,
            false,
            true,
            'text/csv; charset=UTF-8'
        );
	}

    /**
     * Function to send a Predefined report in CSV format
     * @param $report Array with the result of the predefined report
     *        execution
     * @param  $timenow TIMESTAMP
     * @return stdClass Object SQL statement result
     */
	public function send_report_predefined($report, $timenow) {
        $starttime = microtime(true);
        $csvfilename = $this->file_temp_name("report",time());
        $csvtimestamp = $timenow;

        if (!file_exists($csvfilename)) {
            $handle = fopen($csvfilename, 'w');
        } else {
            $handle = fopen($csvfilename, 'a');
        }
        foreach ($report as $row) {
            $this->write_csv_row($handle, $row);
        }
        //echo print_r($handle);
        if (!empty($handle)) {
            fclose($handle);
        }
        send_file(
            $csvfilename,
            'report.csv',
            'default',
            0,
            false,
            true,
            'text/csv; charset=UTF-8'
        );
    }

	private function csv_filename($report, $timenow) {
	    if ($report->runable == 'manual') {
	        return report_customsql_ptt_temp_cvs_name($report->id, $timenow);

	    } else if ($report->singlerow) {
	        return report_customsql_ptt_accumulating_cvs_name($report->id);

	    } else {
	        list($timestart) = report_customsql_ptt_get_starts($report, $timenow);
	        return report_customsql_ptt_scheduled_cvs_name($report->id, $timestart);
	    }
	}

	private function file_temp_name($reportname, $time) {
   		global $CFG;
    	$path = 'report_customsql/temp/CSV/'.$reportname;
    	make_upload_directory($path);
    	return $CFG->dataroot.'/'.$path.'/'.$time.'.csv';
	}

	private function write_csv_row($handle, $row) {
		fwrite($handle,implode(',', $row) ."\r\n");
	}

}