<?php

class Connection {

    public function __construct($host, $login, $password, $dir) {
        $this->ftp_host = $host;
        $this->ftp_login = $login;
        $this->ftp_password = $password;
        $this->ftp_dir = $dir;
    }

    public function connectAndUploadFile($path_to_output_xlsx_file)
    {
        $conn_id = ftp_connect($this->ftp_host);
        $login_result = ftp_login($conn_id, $this->ftp_login, $this->ftp_password);
        // проверяем подключение
        if ((!$conn_id) || (!$login_result)) {
            echo "Connected to $this->ftp_host , for user: $this->ftp_login";
        } else {
            return $this->uploadFile($conn_id, $path_to_output_xlsx_file);
        }

        return $this;
    }

    private function uploadFile($conn_id, $path_to_output_xlsx_file)
    {
        $passive = true;
        ftp_pasv($conn_id, $passive);
        // загружаем файл
        $upload = ftp_put($conn_id, $this->ftp_dir . '/' . $path_to_output_xlsx_file, $path_to_output_xlsx_file, FTP_BINARY);
        // проверяем статус загрузки
        set_time_limit(300);

        if (!$upload) {
            echo "Error: FTP upload has failed!";
        } else {
            echo "Good: Uploaded $path_to_output_xlsx_file to $this->ftp_host";
        }

        ftp_close($conn_id);
    }
}