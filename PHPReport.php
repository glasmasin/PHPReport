<?php

/**
 * PHPReport
 * Library for generating reports from PHP
 * Copyright (c) 2014 PHPReport
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 *
 * @package PHPReport
 * @author Vernes Šiljegović
 * @author Tom Horwood
 * @copyright  Copyright (c) 2014 PHPReport
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version 1.2, 2014-11-21
 */

/**
 * PHPExcel
 *
 * @copyright  Copyright (c) 2006 - 2011 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPReport
{
    //report template
    private $_templateDir;
    private $_template;
    private $_usingTemplate;
    //internal collections of data
    private $_data = array();
    private $_search = array();
    private $_replace = array();
    private $_group = array();
    private $_lastColumn = 'A';
    private $_lastRow = 1;
    //parameters
    private $_renderHeading = false;
    private $_useStripRows = false;
    private $_headingText;
    private $_subheadingText = array();
    private $_noResultText;
    //styling
    private $_headerStyleArray = array(
        'font' => array(
            'bold' => true,
            'color' => array(
                'rgb' => 'FFFFFF'
            )
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array(
                'rgb' => '4E5A7A'
            )
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        )
    );
    private $_footerStyleArray = array(
        'font' => array(
            'bold' => true,
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array(
                'rgb' => 'E4E8F3',
            )
        )
    );
    private $_headerGroupStyleArray = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT
        ),
        'font' => array(
            'bold' => true
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array(
                'rgb' => '8DB4E3'
            )
        )
    );
    private $_footerGroupStyleArray = array(
        'font' => array(
            'bold' => true
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array(
                'rgb' => 'C5D9F1'
            )
        )
    );
    private $_noResultStyleArray = array(
        'borders' => array(
            'outline' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN
            )
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        ),
        'font' => array(
            'bold' => true
        ),
        'fill' => array(
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'startcolor' => array(
                'rgb' => 'FFEBA5'
            )
        )
    );
    private $_headingStyleArray = array(
        'font' => array(
            'bold' => true,
            'color' => array(
                'rgb' => '4E5A7A'
            ),
            'size' => '24'
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        )
    );
    private $_subheadingStyleArray = array(
        'font' => array(
            'bold' => true,
            'color' => array(
                'rgb' => '000000'
            ),
            'size' => '18'
        ),
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        )
    );
    //PHPExcel objects
    private $objReader;

    /**
     * @var PHPExcel
     */
    private $objPHPExcel;

    /**
     * @var PHPExcel_Worksheet
     */
    private $objWorksheet;

    /**
     * @var PHPExcel_Writer_IWriter
     */
    private $objWriter;

    /**
     * Creates new report with some configuration parameters
     * @param array $config
     */
    public function __construct($config = array())
    {
        $this->setConfig($config);
        $this->init();
    }

    /**
     * Uses configuration array to adjust report parameters
     * @param array $config
     */
    public function setConfig($config)
    {
        if (!is_array($config)) {
            throw new Exception('Unable to use non-array configuration');
        }

        foreach ($config as $key => $value) {
            $_key = '_' . $key;
            $this->$_key = $value;
        }
    }

    /**
     * Initializes internal objects
     */
    private function init()
    {
        if ($this->_template != '') {
            $this->loadTemplate();
        } else {
            $this->createTemplate();
        }
    }

    /**
     * Loads Excel file as a template for report
     */
    public function loadTemplate($template = '')
    {
        if ($template != '') {
            $this->_template = $template;
        }

        if (!is_file($this->_templateDir . $this->_template)) {
            throw new Exception('Unable to load template file: ' . $this->_templateDir . $this->_template);
        }

        //identify type of template file
        $inputFileType = PHPExcel_IOFactory::identify($this->_templateDir . $this->_template);
        //TODO: better control of allowed input types
        //load template file into PHPExcel objects
        $this->objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $this->objPHPExcel = $this->objReader->load($this->_templateDir . $this->_template);
        $this->objWorksheet = $this->objPHPExcel->getActiveSheet();

        $this->_usingTemplate = true;
    }

    /**
     * Creates PHPExcel object and template for report
     */
    private function createTemplate()
    {
        $this->objPHPExcel = new PHPExcel();
        $this->objPHPExcel->setActiveSheetIndex(0);
        $this->objWorksheet = $this->objPHPExcel->getActiveSheet();
        //TODO: other parameters
        $this->objWorksheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT);
        $this->objWorksheet->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
        $this->objWorksheet->getPageSetup()->setHorizontalCentered(true);
        $this->objWorksheet->getPageSetup()->setVerticalCentered(false);

        $this->_usingTemplate = false;
    }

    protected function insertRows($beforeRow=1, $numRows=1, $force=false)
    {
        if ($force || $this->objWorksheet->getHighestDataRow() + 1 < $beforeRow) {
            $this->objWorksheet->insertNewRowBefore($beforeRow, $numRows);
        }
    }
    /**
     * Takes an array of all the data for report
     *
     * @param array $dataCollection Associative array with data for report
     *                              or an array of such arrays
     *                              id - unique identifier of data group
     *                              data - Single array of data
     */
    public function load($dataCollection)
    {
        if (!is_array($dataCollection)) {
            throw new Exception("Could not load a non-array data!");
        }

        //clear current data
        $this->clearData();

        //check if it is a single array of data
        if (isset($dataCollection['data'])) {
            $this->addData($dataCollection);
        } else {
            //it's an array of arrays of data, add all
            foreach ($dataCollection as $data) {
                $this->addData($data);
            }
        }
    }

    /**
     * Takes an array of all the data for report
     *
     * @param array $data Associative array with two elements
     *                    id - unique identifier of data group
     *                    data - Single array of data
     */
    private function addData($data)
    {
        if (!is_array($data)) {
            throw new Exception("Could not load a non-array data!");
        }
        if (!isset($data['id'])) {
            throw new Exception("Every array of data needs an 'id'!");
        }
        if (!isset($data['data'])) {
            throw new Exception("Loaded array needs an element 'data'!");
        }

        $this->_data[] = $data;
    }

    /**
     * Clears internal collection of data
     */
    private function clearData()
    {
        $this->_data = array();
    }

    /**
     * Creates a new report based on loaded data
     */
    public function createReport()
    {
        foreach ($this->_data as $data) {
            //$data must have id and data elements
            //$data may also have config, header, footer, group

            $id = $data['id'];
            $format = isset($data['format']) ? $data['format'] : array();
            $config = isset($data['config']) ? $data['config'] : array();
            $group = isset($data['group']) ? $data['group'] : array();

            $configHeader = isset($config['header']) ? $config['header'] : $config;
            $configData = isset($config['data']) ? $config['data'] : $config;
            $configFooter = isset($config['footer']) ? $config['footer'] : $config;

            $config = array(
                'header' => $configHeader,
                'data' => $configData,
                'footer' => $configFooter
            );

            //set the group
            $this->_group = $group;

            $loadCollection = array();

            $row = $this->objWorksheet->getHighestRow();
            if ($row > 1) {
                $row++;
            }

            //form the header for data
            if (isset($data['header'])) {
                $loadCollection[] = $this->prepareDataHeader($data, $id, $row, $config);

                //move to next row for data
                $row++;
            }

            //form the data repeating row
            $dataId = 'DATA_' . $id;
            $colIndex = -1;

            //form the template row
            if (count($data['data']) > 0) {
                //we just need first row of data, to see array keys
                $singleDataRow = $data['data'][0];
                foreach (array_keys($singleDataRow) as $key) {
                    $colIndex++;
                    $tag = "{" . $dataId . ":" . $key . "}";
                    $this->objWorksheet->setCellValueByColumnAndRow($colIndex, $row, $tag);
                    if (isset($config['data'][$key]['align'])) {
                        $this->objWorksheet->getStyleByColumnAndRow($colIndex, $row)->getAlignment()
                                ->setHorizontal($config['data'][$key]['align']);
                    }
                }
            }

            //add this row to collection for load but with repeating
            $loadCollection[] = array('id' => $dataId, 'data' => $data['data'], 'repeat' => true, 'format' => $format);
            $this->enableStripRows();

            //form the footer row for data if needed
            if (isset($data['footer'])) {
                $row++;
                $loadCollection[] = $this->prepareDataFooter($data, $id, $row, $config, $format);
            }

            $this->load($loadCollection);
            $this->generateReport();
        }
    }

    /**
     * Formatting
     * @param  type $data
     * @param  type $id
     * @param  type $row
     * @param  type $config
     * @return type [String, String[]] The header id and data array
     */
    private function prepareDataHeader($data, $id, $row, $config)
    {
        $colIndex = -1;
        $headerId = 'HEADER_' . $id;
        foreach (array_keys($data['header']) as $key) {
            $colIndex++;
            $tag = "{" . $headerId . ":" . $key . "}";
            $this->objWorksheet->setCellValueByColumnAndRow($colIndex, $row, $tag);
            if (isset($config['header'][$key]['width'])) {
                $this->objWorksheet->getColumnDimensionByColumn($colIndex)
                        ->setWidth(pixel2unit($config['header'][$key]['width']));
            }
            if (isset($config['header'][$key]['align'])) {
                $this->objWorksheet->getStyleByColumnAndRow($colIndex, $row)->getAlignment()
                        ->setHorizontal($config['header'][$key]['align']);
            }
        }

        if ($colIndex > -1) {
            $rangeStart = PHPExcel_Cell::stringFromColumnIndex(0) . $row;
            $rangeEnd = PHPExcel_Cell::stringFromColumnIndex($colIndex) . $row;
            $this->objWorksheet->getStyle("$rangeStart:$rangeEnd")->applyFromArray($this->_headerStyleArray);
        }

        return array('id' => $headerId, 'data' => $data['header']);
    }

    private function prepareDataFooter($data, $id, $row, $config, $format)
    {
        $footerId = 'FOOTER_' . $id;
        $colIndex = -1;

        foreach (array_keys($data['footer']) as $key) {
            $colIndex++;
            $tag = "{" . $footerId . ":" . $key . "}";
            $this->objWorksheet->setCellValueByColumnAndRow($colIndex, $row, $tag);
            if (isset($config['footer'][$key]['align'])) {
                $this->objWorksheet->getStyleByColumnAndRow($colIndex, $row)->getAlignment()
                        ->setHorizontal($config['footer'][$key]['align']);
            }
        }
        if ($colIndex > -1) {
            $rangeStart = PHPExcel_Cell::stringFromColumnIndex(0) . $row;
            $rangeEnd = PHPExcel_Cell::stringFromColumnIndex($colIndex) . $row;
            $this->objWorksheet->getStyle("$rangeStart:$rangeEnd")->applyFromArray($this->_footerStyleArray);
        }

        return array('id' => $footerId, 'data' => $data['footer'], 'format' => $format);
    }

    /**
     * Generates report based on loaded data
     */
    public function generateReport()
    {
        $this->_lastColumn = $this->objWorksheet->getHighestColumn(); //TODO: better detection
        $this->_lastRow = $this->objWorksheet->getHighestRow();
        foreach ($this->_data as $data) {
            if ($this->_template != '') {
                $this->_group = isset($data['group']) ? $data['group'] : array();
            }
            if (isset($data['repeat']) && $data['repeat'] == true) {
                //Repeating data
                $foundTags = false;
                $repeatRange = '';
                $firstRow = '';
                $lastRow = '';

                $firstCol = 'A'; //TODO: better detection
                $lastCol = $this->_lastColumn;

                //scan the template
                //search for repeating part
                foreach ($this->objWorksheet->getRowIterator() as $row) {
                    $cellIterator = $row->getCellIterator();
                    $rowIndex = $row->getRowIndex();

                    //find the repeating range (one or more rows)
                    list($foundTags, $lastRow, $firstRow) = $this->findRepeatingSection($cellIterator, $data, $foundTags, $rowIndex, $firstRow, $lastRow);
                }

                //form the repeating range
                if ($foundTags) {
                    $repeatRange = $firstCol . $firstRow . ":" . $lastCol . $lastRow;
                }

                //check if this is the last row
                if ($foundTags && $lastRow == $this->_lastRow) {
                    $data['last'] = true;
                }

                //set initial format data
                if (!isset($data['format'])) {
                    $data['format'] = array();
                }

                //set default step as 1
                if (!isset($data['step'])) {
                    $data['step'] = 1;
                }

                //check if data is an array
                if (is_array($data['data'])) {
                    $this->generateRepeatingRowsFromArray($foundTags, $data, $repeatRange, $firstRow, $lastRow, $firstCol, $lastCol);
                } else {
                    //TODO
                    //maybe an SQL query?
                    //needs to be database agnostic
                }
            } else {
                //non-repeating data
                //check for additional formating
                if (!isset($data['format'])) {
                    $data['format'] = array();
                }

                //check if data is an array or mybe a SQL query
                if (is_array($data['data'])) {
                    //array of data
                    $this->generateSingleRow($data);
                } else {
                    //TODO
                    //maybe an SQL query?
                    //needs to be database agnostic
                }
            }
        }

        //call the replacing function
        $this->searchAndReplace();

        //generate subheadings if any have been added
        if (count($this->_subheadingText) > 0) {
            $this->generateSubHeadings();
        }
        //generate heading if heading text is set
        if ($this->_headingText != '') {
            $this->generateHeading();
        }
    }

    private function findRepeatingSection($cellIterator, $data, $foundTags, $rowIndex, $firstRow, $lastRow)
    {
        foreach ($cellIterator as $cell) {
            $cellval = trim($cell->getValue());
            //see if the cell has something for replacing
            $matches = null;
            if (preg_match_all("/\{" . $data['id'] . ":(\w*|#\+?-?(\d*)?)\}/", $cellval, $matches)) {
                //this cell has replacement tags
                if (!$foundTags) {
                    $foundTags = true;
                }
                //remember the first and the last row
                if ($rowIndex != $firstRow) {
                    $lastRow = $rowIndex;
                }
                if ($firstRow == '') {
                    $firstRow = $rowIndex;
                }
            }
        }
        return array($foundTags, $lastRow, $firstRow);
    }

    private function generateRepeatingRowsFromArray($foundTags, $data, $repeatRange, $firstRow, $lastRow, $firstCol, $lastCol)
    {
        //every element is an array with data for all the columns
        if ($foundTags) {
            //insert repeating rows, as many as needed
            //check if grouping is defined
            $templateArray = $this->objWorksheet->rangeToArray($repeatRange, null, true, true, true);
            if (count($this->_group)) {
                $this->generateRepeatingRowsWithGrouping($data, $repeatRange);
            } else {
                $this->generateRepeatingRows($data, $repeatRange);
            }
            //remove the template rows
            $this->removeTemplateRows($templateArray, $firstRow, $lastRow);
            //if there is no data
            if (count($data['data']) == 0) {
                $this->addNoResultRow($firstRow, $firstCol, $lastCol);
            }
        }
    }

    /**
     * This removes a template row from the spreadsheet, while preserving conditional rules.
     * It is needed as removing a row from the spreadsheet seems to remove any conditional rules from that
     * row number, rather than reapplying them to the row that now sits in that number.
     * @param type $templateArray
     * @return type
     */
    private function removeTemplateRows($templateArray, $firstRow, $lastRow)
    {
        // Save the conditional styles

        $conditionalsArray = array();
        foreach ($templateArray as $rowKey => $rowData) {   //$rowKey is like 9,10,11, ...
            $conditionalsArray[$rowKey] = array();
            foreach (array_keys($rowData) as $col) {    //$col is like A, B, C, ...
                //copy cell styles
                $cellStyle = $this->objWorksheet->getStyle($col . $rowKey);
                $conditionalStyle = $cellStyle->getConditionalStyles();
                $conditionalsArray[$rowKey][$col] = $conditionalStyle;
            }
        }

        $this->objWorksheet->removeRow($firstRow, 1 + $lastRow-$firstRow);

        // Reapply the conditional styles
        foreach ($conditionalsArray as $rowKey => $rowData) {   //$rowKey is like 9,10,11, ...
            foreach ($rowData as $col => $conditionalStyle) {    //$col is like A, B, C, ...
                //copy cell styles
                $this->objWorksheet->getStyle($col . $rowKey)->setConditionalStyles($conditionalStyle);
            }
        }
    }

    /**
     * Generates single non-repeating row of data
     * @param array $data
     */
    private function generateSingleRow(& $data)
    {
        $id = $data['id'];
        $format = $data['format'];
        foreach ($data['data'] as $key => $value) {
            $search = "{" . $id . ":" . $key . "}";
            $this->_search[] = $search;

            //if it needs formating
            if (isset($format[$key])) {
                foreach ($format[$key] as $ftype => $f) {
                    $value = $this->formatValue($value, $ftype, $f);
                }
            }
            $this->_replace[] = $value;
        }
    }

    /**
     * Generates repeating rows of data with some template range
     * @param array $data
     * @param string $repeatRange
     */
    private function generateRepeatingRows(& $data, $repeatRange)
    {
        $rowCounter = 0;
        $templateArray['values'] = $this->objWorksheet->rangeToArray($repeatRange, null, true, true, true);
        $templateArray['conditionalStyles'] = array();
        $templateArray['xfIndices'] = array();

        foreach ($templateArray['values'] as $rowIndex => $rowCells) {
            $templateArray['styles'][$rowIndex] = array();
            foreach (array_keys($rowCells) as $colIndex) {
                $cellRef = $colIndex . $rowIndex;
                $templateArray['conditionalStyles'][$rowIndex][$colIndex] = $this->objWorksheet->getConditionalStyles($cellRef);
                $templateArray['xfIndices'][$rowIndex][$colIndex] = $this->objWorksheet->getCell($cellRef)->getXfIndex();
            }
        }
        //insert repeating rows but first check for minimum number of rows
        if (isset($data['minRows'])) {
            $minRows = (int) $data['minRows'];
        } else {
            $minRows = 0;
        }

        //is this the last data
        if (isset($data['last'])) {
            $last = $data['last'];
        } else {
            $last = false;
        }

        $templateKeys = array_keys($templateArray['values']);
        $lastRowFoundAt = end($templateKeys);
        $rowsFound = count($templateArray['values']);

        $mergeCells = $this->objWorksheet->getMergeCells();
        $needMerge = array();

        foreach ($mergeCells as $mergeCell) {
            if ($this->isSubrange($mergeCell, $repeatRange)) {
                //contains merged cells, save for later
                $needMerge[] = $mergeCell;
            }
        }

        //check if any new rows need to bi inserted
        $dataRows = count($data['data']);
        if ($minRows < $dataRows) {
            $this->insertRows($lastRowFoundAt + 1, $rowsFound * ($dataRows - $minRows));
        }

        //check all the data
        foreach ($data['data'] as $value) {
            $rowCounter++;
            $skip = $rowCounter * $rowsFound;

            //copy merge definitions
            foreach ($needMerge as $nm) {
                $nm = PHPExcel_Cell::rangeBoundaries($nm);
                $rangeStart = PHPExcel_Cell::stringFromColumnIndex($nm[0][0] - 1) . ($nm[0][1] + $skip);
                $rangeEnd = PHPExcel_Cell::stringFromColumnIndex($nm[1][0] - 1) . ($nm[1][1] + $skip);
                $newMerge = "$rangeStart:$rangeEnd";

                $this->objWorksheet->mergeCells($newMerge);
            }

            //generate row of data
            $this->generateSingleRepeatingRow($value, $templateArray, $rowCounter, $skip, $data['id'], $data['format'], $data['step']);
        }
        //remove merge on template, BUG fix
        foreach ($needMerge as $nm) {
            $this->objWorksheet->unmergeCells($nm);
        }
    }

    /**
     * Generates repeating rows of data with some template range but also with grouping
     * @param array  $data
     * @param string $repeatRange
     */
    private function generateRepeatingRowsWithGrouping(& $data, $repeatRange)
    {
        $rowCounter = 0;
        $groupCounter = 0;
        $footerCount = 0;
        $templateArray['values'] = $this->objWorksheet->rangeToArray($repeatRange, null, true, true, true);
        $templateArray['conditionalStyles'] = array();
        $templateArray['xfIndices'] = array();

        foreach ($templateArray['values'] as $rowIndex => $rowCells) {
            $templateArray['styles'][$rowIndex] = array();
            foreach (array_keys($rowCells) as $colIndex) {
                $cellRef = $colIndex . $rowIndex;
                $templateArray['conditionalStyles'][$rowIndex][$colIndex] = $this->objWorksheet->getConditionalStyles($cellRef);
                $templateArray['xfIndices'][$rowIndex][$colIndex] = $this->objWorksheet->getCell($cellRef)->getXfIndex();
            }
        }
        //insert repeating rows but first check for minimum number of rows

        if (isset($data['minRows'])) {
            $minRows = (int) $data['minRows'];
        } else {
            $minRows = 0;
        }

        $templateKeys = array_keys($templateArray['values']);
        $firstRowFoundAt = reset($templateKeys);
        $rowsFound = count($templateArray['values']);

        $headerRowCount = count($this->_group['caption']);
        $detailRowCount = count($templateArray['values']) * count($data['data']);
        $footerRowCount = isset($this->_group['summary']) ? count($this->_group['summary']) : 0;
        $rowsToInsert = $headerRowCount + $detailRowCount + $footerRowCount;
        $this->insertRows($firstRowFoundAt + 1, $rowsToInsert);

        list($rangeStart, $rangeEnd) = PHPExcel_Cell::rangeBoundaries($repeatRange);
        $firstCol = PHPExcel_Cell::stringFromColumnIndex($rangeStart[0] - 1);
        $lastCol = PHPExcel_Cell::stringFromColumnIndex($rangeEnd[0] - 1);

        $mergeCells = $this->objWorksheet->getMergeCells();
        $needMerge = array();
        foreach ($mergeCells as $mergeCell) {
            if ($this->isSubrange($mergeCell, $repeatRange)) {
                //contains merged cells, save for later
                $needMerge[] = $mergeCell;
            }
        }

        //group array should have header, rows and summary elements
        foreach ($this->_group['rows'] as $name => $rows) {
            $groupCounter++;
            $caption = $this->_group['caption'][$name];
            $newRowIndex = $firstRowFoundAt + $rowCounter * $rowsFound + $footerCount * $rowsFound + $groupCounter;
            //insert header for the group
            $this->objWorksheet->setCellValue($firstCol . $newRowIndex, $caption);
            $this->objWorksheet->mergeCells($firstCol . $newRowIndex . ":" . $lastCol . $newRowIndex);

            //add style for the header

            $this->objWorksheet->getStyle($firstCol . $newRowIndex)->applyFromArray($this->_headerGroupStyleArray);

            //add data for the group
            foreach ($rows as $row) {
                $rowData = $data['data'][$row];
                $rowCounter++;
                $skip = $rowCounter * $rowsFound + $footerCount * $rowsFound + $groupCounter;
                $newRowIndex = $firstRowFoundAt + $skip;

                //copy merge definitions
                foreach ($needMerge as $nm) {
                    $nm = PHPExcel_Cell::rangeBoundaries($nm);
                    $rangeStart = PHPExcel_Cell::stringFromColumnIndex($nm[0][0] - 1) . ($nm[0][1] + $skip);
                    $rangeEnd = PHPExcel_Cell::stringFromColumnIndex($nm[1][0] - 1) . ($nm[1][1] + $skip);
                    $newMerge = "$rangeStart:$rangeEnd";

                    $this->objWorksheet->mergeCells($newMerge);
                }

                //generate row of data
                $this->generateSingleRepeatingRow($rowData, $templateArray, $rowCounter, $skip, $data['id'], $data['format'], $data['step']);
            }

            //include the footer if defined
            if (isset($this->_group['summary']) && isset($this->_group['summary'][$name])) {
                $footerCount++;
                $skip = $groupCounter + $rowCounter * $rowsFound + $footerCount * $rowsFound;
                $newRowIndex = $firstRowFoundAt + $skip;

                $this->generateSingleRepeatingRow($this->_group['summary'][$name], $templateArray, '', $skip, $data['id'], $data['format'], $data['step']);
                //add style for the footer

                $this->objWorksheet->getStyle($firstCol . $newRowIndex . ":" . $lastCol . $newRowIndex)->applyFromArray($this->_footerGroupStyleArray);
            }

            //remove merge on template, BUG fix
            foreach ($needMerge as $nm) {
                $this->objWorksheet->unmergeCells($nm);
            }
        }
    }

    /**
     * Generates single row for repeating data
     * @param array  $values
     * @param array  $template
     * @param int    $rowCounter
     * @param int    $skip
     * @param string $id
     * @param array  $format
     * @param int    $step
     */
    private function generateSingleRepeatingRow(& $values, & $template, $rowCounter, $skip, $id, $format, $step)
    {
        foreach ($template['values'] as $rowKey => $templateRow) {
            foreach ($templateRow as $col => $templateCellContent) {
                //$col is like A, B, C, ...
                //$rowKey is like 9,10,11, ...
                //$tag can have many replacement tags, e.g. "{item:item_id} --- {item:item_code}"
                $matches = null;
                if (preg_match_all("/\{" . $id . ":(\w*|#\+?-?(\d*)?)\}/", $templateCellContent, $matches)) {
                    $templateCellContent = $this->generateCellValue($templateCellContent, $values, $rowCounter, $id, $format, $step, $matches);
                }
                $newCellAddress = $col . ($rowKey + $skip);
                $this->objWorksheet->setCellValue($newCellAddress, $templateCellContent);

                //copy cell styles
                $conditionalStyles = $template['conditionalStyles'][$rowKey][$col];
                if ($conditionalStyles) {
                    $this->objWorksheet->setConditionalStyles($newCellAddress, $conditionalStyles);
                }
                $xfIndex = $template['xfIndices'][$rowKey][$col];
                $this->objWorksheet->getCell($newCellAddress)->setXfIndex($xfIndex);

                //strip rows if requested
                if ($this->_useStripRows && $rowCounter % 2) {
                    $this->objWorksheet->getStyle($newCellAddress)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                    $this->objWorksheet->getStyle($newCellAddress)->getFill()->getStartColor()->setRGB('F2F2F2');
                }
            }
        }
    }

    private function generateCellValue($cellContent, $value, $rowCounter, $id, $format, $step, $matches)
    {
        $matchKeys = $matches[1]; //array with only the key names, e.g. 'item_id'
        $replaceTags = array();
        $replaceValues = array();

        foreach ($matchKeys as $mkey) {
            $replaceTags[] = "{" . $id . ":" . $mkey . "}";
            if (strpos($mkey, "#") === 0) {
                //this is a counter (optional offset)
                $offset = explode("+", $mkey);
                if (count($offset) > 1) {
                    $offset = $offset[1];
                } else {
                    $offset = 0;
                }

                $rValue = ($rowCounter - 1) * $step + 1 + (int) $offset;
            } elseif (key_exists($mkey, $value)) {
                $rValue = $this->generateTagValue($format, $mkey, $value);
            } else {
                $rValue = $mkey;
            }

            //add to replace array
            $replaceValues[] = $rValue;
        }
        //replace all the values in this cell
        return str_replace($replaceTags, $replaceValues, $cellContent);
    }

    private function generateTagValue($format, $mkey, $value)
    {
        //format if needed
        if (isset($format) && isset($format[$mkey])) {
            foreach ($format[$mkey] as $ftype => $f) {
                $rValue = $this->formatValue($value[$mkey], $ftype, $f);
            }
        } else {
            //without additional formating
            $rValue = $value[$mkey];
        }

        return $rValue;
    }

    /**
     * Check and apply various formating
     */

    /**
     * Applies various formatings
     * Type can be datetime or number
     * @param mixed $value
     * @param string $type
     * @param mixed $format
     */
    protected function formatValue($value, $type, $format)
    {
        if ($type == 'datetime') {
            //format can only be string
            if (is_string($format)) {
                $value = date($format, strtotime($value));
            }
        } elseif ($type == 'number') {
            //format must be an array
            if (is_array($format)) {
                //set the defaults
                if (!isset($format['prefix'])) {
                    $format['prefix'] = '';
                }
                if (!isset($format['decimals'])) {
                    $format['decimals'] = 0;
                }
                if (!isset($format['decPoint'])) {
                    $format['decPoint'] = '.';
                }
                if (!isset($format['thousandsSep'])) {
                    $format['thousandsSep'] = ',';
                }
                if (!isset($format['sufix'])) {
                    $format['sufix'] = '';
                }
                $value = $format['prefix'] . number_format($value, $format['decimals'], $format['decPoint'], $format['thousandsSep']) . $format['sufix'];
            }
        }

        return $value;
    }

    /**
     * Replaces all the cells with real data
     */
    private function searchAndReplace()
    {
        foreach ($this->objWorksheet->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            foreach ($cellIterator as $cell) {
                $cell->setValue(str_replace($this->_search, $this->_replace, $cell->getValue()));
            }
        }
    }

    /**
     * Adda a row for repeating data when there is no results
     * @param int $rowIndex
     * @param string $colMin
     * @param string $colMax
     */
    private function addNoResultRow($rowIndex, $colMin, $colMax)
    {
        //merge as required
        $this->objWorksheet->mergeCells($colMin . $rowIndex . ":" . $colMax . $rowIndex);

        //insert text

        $this->objWorksheet->setCellValue($colMin . $rowIndex, $this->_noResultText);

        $this->objWorksheet->getStyle($colMin . $rowIndex . ":" . $colMax . $rowIndex)
                ->applyFromArray($this->_noResultStyleArray);
    }

    /**
     * Generates subheading titles of the report
     */
    private function generateSubHeadings()
    {
        //get current dimensions
        $highestColumn = $this->objWorksheet->getHighestColumn(); // e.g 'F'
        //insert rows on top
        for ($i = count($this->_subheadingText)-1; $i >= 0 ; $i--) {
            $text = $this->_subheadingText[$i];

            $this->insertRows(1, 1, true);

            //merge cells
            $this->objWorksheet->mergeCells("A1:" . $highestColumn . "1");

            //set the text for header
            $this->objWorksheet->setCellValue("A1", $text);
            $this->objWorksheet->getStyle('A1')->getAlignment()->setWrapText(true);
            $this->objWorksheet->getRowDimension('1')->setRowHeight(24);

            //Apply style
            $this->objWorksheet->getStyle("A1")->applyFromArray($this->_subheadingStyleArray);
        }
    }

    /**
     * Generates heading title of the report
     */
    private function generateHeading()
    {
        //get current dimensions
        $highestColumn = $this->objWorksheet->getHighestColumn(); // e.g 'F'
        //insert row on top
        $this->insertRows(1, 1, true);

        //merge cells
        $this->objWorksheet->mergeCells("A1:" . $highestColumn . "1");

        //set the text for header
        $this->objWorksheet->setCellValue("A1", $this->_headingText);
        $this->objWorksheet->getStyle('A1')->getAlignment()->setWrapText(true);
        $this->objWorksheet->getRowDimension('1')->setRowHeight(48);

        //Apply style
        $this->objWorksheet->getStyle("A1")->applyFromArray($this->_headingStyleArray);
    }

    /**
     * Renders report as specified output file
     * @param string $type
     * @param string $filename
     */
    public function render($type = 'html', $filename = '')
    {
        //create or generate report
        if ($this->_usingTemplate) {
            $this->generateReport();
        } else {
            $this->createReport();
        }

        if ($type == '') {
            $type = "html";
        }

        if ($filename == '') {
            $filename = "Report " . date("Y-m-d");
        } else {
            $filename = strftime($filename);
        }
        //http://strftime.net/


        if (strtolower($type) == 'html') {
            return $this->renderHtml();
        } elseif (strtolower($type) == 'excel') {
            return $this->renderXlsx($filename);
        } elseif (strtolower($type) == 'excel2003') {
            return $this->renderXls($filename);
        } elseif (strtolower($type) == 'pdf') {
            return $this->renderPdf($filename);
        } else {
            return "Error: unsupported export type!";
        } //TODO: better error handling
    }

    /**
     * Renders report as a HTML output
     */
    protected function renderHtml()
    {
        $this->objWriter = new PHPExcel_Writer_HTML($this->objPHPExcel);
        //$this->objWriter->setPreCalculateFormulas(false);
        // Generate HTML
        $html = '';
        $html .= $this->objWriter->generateHTMLHeader(true);
        $html .= $this->objWriter->generateSheetData();
        $html .= $this->objWriter->generateHTMLFooter();
        $html .= '';
        $this->objPHPExcel->disconnectWorkSheets();
        unset($this->objWriter);
        unset($this->objWorksheet);
        unset($this->objReader);
        unset($this->objPHPExcel);

        return $html;
    }

    /**
     * Renders report as a XLSX file
     */
    protected function renderXlsx($filename)
    {
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
        header('Cache-Control: max-age=0');

        $this->objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');

        $this->objWriter->save('php://output');
        unset($this->objWriter);
        unset($this->objWorksheet);
        unset($this->objReader);
        unset($this->objPHPExcel);
        exit();
    }

    /**
     * Renders report as a XLS file
     */
    protected function renderXls($filename)
    {
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $filename . '.xls"');
        header('Cache-Control: max-age=0');

        $this->objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');

        $this->objWriter->save('php://output');
        unset($this->objWriter);
        unset($this->objWorksheet);
        unset($this->objReader);
        unset($this->objPHPExcel);
        exit();
    }

    /**
     * Renders report as a PDF file
     */
    protected function renderPdf($filename)
    {
        header('Content-Type: application/vnd.pdf');
        header('Content-Disposition: attachment;filename="' . $filename . '.pdf"');
        header('Cache-Control: max-age=0');

        $this->objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'PDF');

        $this->objWriter->save('php://output');
        exit();
    }

    /**
     * Helper function for checking subranges of a range
     */
    public function isSubrange($subRange, $range)
    {
        list($rangeStart, $rangeEnd) = PHPExcel_Cell::rangeBoundaries($range);
        list($subrangeStart, $subrangeEnd) = PHPExcel_Cell::rangeBoundaries($subRange);

        return (($subrangeStart[0] >= $rangeStart[0]) && ($subrangeStart[1] >= $rangeStart[1]) && ($subrangeEnd[0] <= $rangeEnd[0]) && ($subrangeEnd[1] <= $rangeEnd[1]));
    }

    /**
     * Enabling strip rows
     */
    public function enableStripRows()
    {
        $this->_useStripRows = true;
    }

    /**
     * Sets title of the report header
     */
    public function setHeading($h)
    {
        $this->_headingText = $h;
    }

    /**
     * Sets title of the report header
     */
    public function addSubHeading($h)
    {
        $this->_subheadingText[] = $h;
    }
}

/**
 * converts pixels to excel units
 * @param float $p
 * @return float
 */
function pixel2unit($p)
{
    return ($p - 5) / 7;
}
