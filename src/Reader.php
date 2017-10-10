<?php

namespace SimplePHPExcelReader;
/**
 * Created by IntelliJ IDEA.
 * User: Martin Jirásek <martin.jirasek@nms.cz>
 * Date: 19.04.2016
 * Time: 11:06
 */
class Reader
{
    /**
     * @var string
     */
    private $filename;

    /**
     * @var \PHPExcel
     */
    private $phpExcel;

    /**
     * @var array
     */
    private $data = [];

    /**
     * Reader constructor.
     * @param string $filename
     */
    public function __construct($filename)
    {
        $this->filename = $filename;
    }

    /**
     * @throws \PHPExcel_Reader_Exception
     */
    private function loadPhpExcel()
    {
        $this->phpExcel = \PHPExcel_IOFactory::createReader(\PHPExcel_IOFactory::identify($this->filename))
            ->load($this->filename);
    }

    /**
     * @param int $fromRow
     * @param bool|false $refresh
     * @return array
     */
    public function read($fromRow = 1, $refresh = false)
    {
        if (!$this->phpExcel || $refresh) {
            $this->loadPhpExcel();
            $this->loadData($fromRow);
        }
        return $this->data;
    }

	/**
	 * @param int $fromRow
	 * @param bool $refresh
	 * @return array
	 */
	public function readAssoc($fromRow = 1, $refresh = false, $omitHeader = false)
	{
		$data = $this->read($fromRow, $refresh);
		$assocData = [];
		$header = $omitHeader ? array_shift($data) : reset($data);
		foreach ($data as $row) {
			$assocData[] = array_combine($header, $row);
		}
		$this->data = $assocData;
		return $this->data;
	}

    /**
     * @param $fromRow
     * @throws \PHPExcel_Exception
     */
    private function loadData($fromRow)
    {
        $sheet = $this->phpExcel->getSheet(0);
        $highestRow = $sheet->getHighestRow();
        $highestColumn = $sheet->getHighestColumn();
        for ($row = $fromRow; $row <= $highestRow; $row++) {
            $data = $sheet->rangeToArray("A$row:$highestColumn" . $row, NULL, TRUE, FALSE);
            $this->data[] = $data[0];
        }
    }
}