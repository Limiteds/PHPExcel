<?php

require_once 'Connection.php';
require_once 'Classes/PHPExcel.php';
require_once 'Classes/PHPExcel/Writer/Excel2007.php';

class XlsExchange
{

    protected $path_to_input_json_file = 'order.json';
    protected $path_to_output_xlsx_file = 'items.xlsx';
    protected $ftp_host = '';
    protected $ftp_login = '';
    protected $ftp_password = '';
    protected $ftp_dir = '';

    private $style_header = [
        // Шрифт
        'font' => [
            'bold' => true,
        ],
        // Выравнивание
        'alignment' => [
            'horizontal' => PHPExcel_STYLE_ALIGNMENT::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_STYLE_ALIGNMENT::VERTICAL_CENTER,
        ],
    ];
    
    private $header = [
        "A1" => "Id",
        "B1" => "ШК",
        "C1" => "Название",
        "D1" => "Кол-во",
        "E1" => "Сумма"
    ];

    private $width_column = [
        "A" => 15,
        "B" => 15,
        "C" => 45
    ];

    public function setInputFile($file_name)
    {
        $this->path_to_input_json_file = $file_name;

        return $this;
    }

    public function setOutputFile($file_name)
    {
        $this->path_to_output_xlsx_file = $file_name;

        return $this;
    }

    public function export()
    {   
        $data = file_get_contents($this->path_to_input_json_file);
        $obj = json_decode($data, true);
        $items = $this->getItemsWithValidateBarCode($obj);
        $xls = $this->createExcelFile($items);

        (new PHPExcel_Writer_Excel2007($xls))
            ->save($this->path_to_output_xlsx_file);

        //Подключение к FTP серверу и отправка файла
        if ($this->checkFTPconfig()) {
           (new connection($this->ftp_host, $this->ftp_login, $this->ftp_password, $this->ftp_dir))
                ->connectAndUploadFile($this->path_to_output_xlsx_file);
        }

        return $this;
    }

    private function getItemsWithValidateBarCode($items)
    {
        $pattern = '#^[0-9]+$#';
        $result = [];
        
        if (is_array($items)) {
            foreach ($items["items"] as $item) {
                if (preg_match($pattern, $item["item"]["barcode"]) && strlen($item["item"]["barcode"]) == 13) {

                    $result[] = $item;
                };
            }
        }
        return $result;
    }

    private function createExcelFile($items)
    {
        
        $xls = new PHPExcel();
        // Устанавливаем индекс активного листа
        $xls->setActiveSheetIndex(0);

        // Получаем активный лист
        $sheet = $xls->getActiveSheet();

        $sheet->getStyle('A1:E1')->applyFromArray($this->style_header);

        // Подписываем лист
        $sheet->setTitle('Продукты');

        foreach ($this->header as $key => $value) {
            $sheet->setCellValue($key, $value);
        }

        foreach ($this->width_column as $key => $value) {
            $sheet->getColumnDimension($key)->setWidth($value);
        }

        if (is_array($items)){
            for ($i = 0; $i < count($items); $i++) {
                $sheet->setCellValueByColumnAndRow(0, $i + 2, $items[$i]["id"]);
                $sheet->getCellByColumnAndRow(1, $i + 2)->setValueExplicit($items[$i]["item"]["barcode"], PHPExcel_Cell_DataType::TYPE_STRING);
                $sheet->setCellValueByColumnAndRow(2, $i + 2, $items[$i]["item"]["name"]);
                $sheet->setCellValueByColumnAndRow(3, $i + 2, $items[$i]["quantity"]);
                $sheet->setCellValueByColumnAndRow(4, $i + 2, $items[$i]["amount"]);
            }
         }

        return $xls;
    }

    private function checkFTPconfig()
    {
        if (empty($this->ftp_host)) return false;
        if (empty($this->ftp_login)) return false;
        if (empty($this->ftp_password)) return false;
        if (empty($this->ftp_dir)) return false;
        return true;
    }

}
