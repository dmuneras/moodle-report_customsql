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

/**
* CustomSqlPttExcel
*
*
* @package    report_customsql
* @subpackage export_format
* @author     Daniel Múnera Sánchez <dmunera119@gmail.com>
*/
interface CustomSqlExportFormat{

	public function send_report_predefined($report, $timenow);
  public function send_report_customsql($id, $csvtimestamp, $params =array());

}