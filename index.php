<?php
require_once "vendor/autoload.php"; // подключение библиотек
require_once "messages.php";
ini_set('display_errors', 0); // 4-6 вывод всех ошибок отключен, указать на 1 для дебага
ini_set('display_startup_errors', 0);
//error_reporting(E_ALL);
ignore_user_abort(true); // при отключении клиента не закрывать скрипт
set_time_limit(0); // отключение лимита работы скрипта
// Библиотеки
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use DigitalStar\vk_api\vk_api; // Основной класс vk_api
use DigitalStar\vk_api\VkApiException; // Обработка ошибок
use Box\Spout\Common\Entity\Row;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

// Данные для подключения
const VK_KEY = ""; // ключ доступа сообщества
const CONFIRM_STR = ""; //секретный ключ vk api
const VERSION = "5.131"; // версия vk api
define('PATH', $_SERVER['DOCUMENT_ROOT']."/"); // директория (если требуется указать поддиректорию, то добавить ее после /)

// Инициализация объекта класса для работы с vk api
try {
    $vk = vk_api::create(VK_KEY, VERSION)->setConfirm(CONFIRM_STR); // авторизация через токен группы/пользователя
    $vk->initVars($id, $message, $payload, $user_id, $type, $data); // получение объекта сообщения
    //$vk->debug(); // дебаг при наличии ошибок с вк апи, появляются ошибки в папке errors
} catch (VkApiException $e) {
    echo $e->getMessage(); // выкидываем ошибку в случае, если токен или секретный ключ неверный.
    exit;
}

/**
 * Некоторые методы для удобной работы
 */
class Functions{

    /**
     *
     * Объект vk api
     *
     * @var object
     */
    private object $vk; // объект класса vk api

    /**
     * Functions constructor.
     * @param $vk - Объект vk api
     */
    function __construct($vk)
    {
        $this->vk = $vk;
    }

    /**
     *
     * Отправляем сообщение отправителю с проверкой на исключение
     *
     * @param $message - сообщение, которое отправляем пользователю
     */
    public function sendReply($message){
        try {
            $this->vk->reply($message);
        } catch (VkApiException $e) {
            echo $e->getMessage(); // выкидываем ошибку в случае если пользователь запретил сообщения от сообещства.
            exit;
        }
    }

    /**
     *
     * Проверяем, сколько параметров отправил пользователь.
     * Если указываем, что разрешенное кол-во параметров 2,
     * то в случае, если их больше или меньше двух, то
     * выкидываем ошибку неверного формата команды.
     *
     * @param $command - массив команд
     * @param $key - разрешенное кол-во параметров
     * @return bool результат проверки
     */
    public function isCorrectCommand($command, $key): bool
    {
        if(count($command) == $key){
            return true;
        }
        else{
            return false;
        }
    }

    /**
     *
     * Проверяем команду на соответствие
     * при выдачи результатов о кураторе с успеваемостью.
     *
     * @param $command
     * @return array
     */
    public function separatorOfCuratorProgress($command): array
    {
        if($command[3] == Messages::EGE or $command[3] == Messages::OGE or $command[3] == Messages::TEN or $command[3] == Messages::CIS){
            if($command[4] == Messages::HIGH or $command[4] == Messages::MID or $command[4] == Messages::LOW){
                return [
                    "name" => "{$command[1]} {$command[2]}",
                    "type" => $command[3],
                    "progress" => $command[4]
                ];
            } else{
                $this->sendReply(Messages::ProgressTypeError);
                exit;
            }
        } else{
            $this->sendReply(Messages::ExamTypeError);
            exit;
        }
    }

    /**
     *
     * Проверяем команду на соответствие
     * при выдачи результатов о кураторе без успеваемости.
     *
     * @param $command
     * @return array
     */
    public function separatorOfCurator($command): array
    {
        if($command[3] == Messages::EGE or $command[3] == Messages::OGE or $command[3] == Messages::TEN or $command[3] == Messages::CIS){
            return [
                "name" => "{$command[1]} {$command[2]}",
                "type" => $command[3]
            ];
        } else{
            $this->sendReply(Messages::ExamTypeError);
            exit;
        }
    }

    /**
     *
     * Скачивание файла от вк
     *
     * @param $data - запрос вк
     * @param $PATH - директория файла
     * @param bool $isReply - нужно ли уведомить о загрузке  файла
     * @param string $allow_ext - разрешенный формат файла
     */
    function downloadFile($data, $PATH, $isReply = true, $allow_ext = Messages::CSV) {
        foreach ($data->object->attachments as $attachment){
            if($attachment->type == Messages::DOC){
                if($attachment->doc->ext == $allow_ext){
                    $URL = $attachment->doc->url;
                }
            }
        }
        if(!isset($URL)){
            $this->sendReply(Messages::FileError.Messages::CSV);
            exit;
        }
        $ReadFile = fopen ($URL, "rb");
        if ($ReadFile) {
            $WriteFile = fopen ($PATH, "wb");
            if ($WriteFile){
                while(!feof($ReadFile)) {
                    fwrite($WriteFile, fread($ReadFile, 4096 ));
                }
                fclose($WriteFile);
            }
            fclose($ReadFile);
        }

        if($isReply)
            $this->sendReply(Messages::FileSuccess);
    }

    /**
     *
     * Разделение строки на элементы
     *
     * @param $string - строка
     * @param $start - начало того, что нужно вытащить
     * @param $end - конец того, что нужно вытащить
     * @return false|string - Если найдена, возвращаем, иначе возвращаем false
     */
    function get_string_between($string, $start, $end){
        $string = ' ' . $string;
        $ini = strpos($string, $start);
        if ($ini == 0) return '';
        $ini += strlen($start);
        $len = strpos($string, $end, $ini) - $ini;
        return substr($string, $ini, $len);
    }

    /**
     *
     * Содержит ли строка только пробелы
     *
     * @param $string - строка
     * @return bool true/false
     */
    function isStringSpace($string) : bool {
        return strlen(preg_replace('/\s+/u','', $string)) == 0;
    }

    /**
     *
     * Группирвока типа просмотров
     *
     * @param $array
     * @return string
     */
    function groupingViews($array): string
    {
        $elementCount = array();
        $result_array = array();
        $result = "(";

        for($i=0; $i < count($array); $i++)
        {
            $key = $array[$i];

            if($elementCount[$key] >= 1)
            {
                $elementCount[$key]++;
            }
            else {
                $elementCount[$key]=1;
            }
        }
        foreach ($elementCount as $item => $key){
            array_push($result_array, "$item - $key");
        }
        $i = 0;
        foreach ($result_array as $item){
            ++$i === count($result_array) ? $result .= $item : $result .= $item.', ';
        }

        return $result.')';
    }

    /**
     *
     * Группирвока типа просмотров со следующей строки
     *
     * @param $array
     * @return string
     */
    function groupingViewsByNextLine($array): string
    {
        $elementCount = array();
        $result_array = array();
        $result = "";

        for($i=0; $i < count($array); $i++)
        {
            $key = $array[$i];

            if($elementCount[$key] >= 1)
            {
                $elementCount[$key]++;
            }
            else {
                $elementCount[$key]=1;
            }
        }
        foreach ($elementCount as $item => $key){
            array_push($result_array, "$item - $key");
        }
        $i = 0;
        foreach ($result_array as $item){
            ++$i === count($result_array) ? $result .= $item : $result .= $item." \n";
        }

        return $result;
    }
}

class DocReader{

    private object $functions;
    private object $vk;
    private $user_id;

    private string $ege = PATH."db/ege.csv"; // файл с учениками курса егэ
    private string $oge = PATH."db/oge.csv"; // файл с учениками курса огэ
    private string $ten_class = PATH."db/ten_class.csv"; // файл с учениками курса 10 класса
    private string $cis = PATH."db/cis.csv"; // файл с учениками курса снг
    private string $tasks = PATH."db/tasks.csv"; // файл с кол-вом занятий по предметам
    private string $dynamic = PATH."dynamic/temp.txt"; // файл со ссылками на учеников, загружается перед обращением
    private string $ege_old = PATH."db/old_progress/ege.csv"; // файл с учениками курса егэ старая
    private string $oge_old = PATH."db/old_progress/oge.csv"; // файл с учениками курса огэ старая
    private string $ten_class_old = PATH."db/old_progress/ten_class.csv"; // файл с учениками курса 10 класса старая
    private string $cis_old = PATH."db/old_progress/cis.csv"; // файл с учениками курса снг старая


    /**
     * DocReader constructor.
     * @param $functions - объект класса функций
     * @param $vk - обхект класса вк
     * @param $user_id - id пользователя
     */
    public function __construct($functions, $vk, $user_id)
    {
        $this->functions = $functions;
        $this->vk = $vk;
        $this->user_id = $user_id;
    }

    /**
     *
     * Получение данных учеников куратора по фильтрам: тип экзамена, успеваемость
     *
     * @param $curator - данные о запросе куратора
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function curatorStudentProgress($curator){
        switch ($curator['type']){
            case Messages::EGE:
                $file = $this->ege;
                break;
            case Messages::OGE:
                $file = $this->oge;
                break;
            case Messages::TEN:
                $file = $this->ten_class;
                break;
            case Messages::CIS:
                $file = $this->cis;
                break;
            default:
                $this->functions->sendReply(Messages::ExamTypeError);
                exit;
        }
        $this->functions->sendReply(Messages::DataProcessing);
        $reader = ReaderEntityFactory::createCSVReader(); // создаем объект класса для чтения файла базы
        $reader->setFieldDelimiter(Messages::COMMA);
        $students = array();
        try {
            $reader->open($file); // открываем нужный файл
            foreach ($reader->getSheetIterator() as $sheet) { // перебираем листы
                foreach ($sheet->getRowIterator() as $row) { // перебираем строки
                    $cells = $row->getCells(); // получаем ячейки в виде массива объектов класса ридера
                    if($cells[5]->getValue() == $curator['name']){ // ищем нужного человека в соотвествии с запросом
                        $temp = $row->toArray();
                        array_push($students, [
                            $temp[2], // "ссылка на ученика",
                            $temp[3], // "Группы",
                            $temp[9], // "Как посмотрел",
                            $temp[13], // "Баллы",
                            $temp[5], // "ФИ куратора",
                            $temp[14], // Задание",
                            $temp[15], // "Предмет"
                            ]); // добавляем в массив ученика, если такой найден в базе
                    }
                }
            }
        } catch (IOException $e) {
            echo $e->getMessage()." ".$e->getLine(); // выкидываем ошибку, если что-то с документом
            exit;
        }

        if(!empty($students)){ // если найдена хоть 1 строчка с учеником
            $subjects = array();
            $tasks = array();

            foreach ($students as $value) { // сортируем сначала по ученикам
                $subjects[$value[0]][] = $value;
            }

            $subjects = array_values($subjects);
            $temp_array = array();
            foreach ($subjects as $subject){ // сортируем по предметам
                foreach ($subject as $item){
                    $temp_array[$item[6]][] = $item;
                }
                $temp_array = array_values($temp_array);
            }
            $subjects = $temp_array;

            unset($reader);
            $reader = ReaderEntityFactory::createCSVReader();
            $reader->setFieldDelimiter(Messages::COMMA);
            try {
                $reader->open($this->tasks);
                for($i = 0; $i < count($subjects); $i++){
                    foreach ($reader->getSheetIterator() as $sheet) {
                        foreach ($sheet->getRowIterator() as $row) {
                            $cells = $row->getCells();
                            if($cells[0]->getValue() == $curator['type'] and $cells[1]->getValue() == $subjects[$i][0][6]){
                                array_push($tasks, ['task' => $subjects[$i][0][6], 'count' => $cells[2]->getValue()]); // составляем массив занятий каждого предмета
                            }
                        }
                    }
                }
            } catch (IOException $e) {
                echo $e->getMessage();
            }

            $tasks = array_unique($tasks, SORT_REGULAR); // отсекаем повторяющиеся значения учеников

            $users = array();
            for($i = 0; $i < count($subjects); $i++){
                $views = array();
                foreach ($tasks as $task){
                    if($task['task'] == $subjects[$i][0][6]){ // идем по массиву
                        $viewed_percent = count($subjects[$i]) / (int)$task['count'] * 100; // кол-во в процентах просмотренных занятий
                        $avg = 0; // средний балл
                        $homeworks = 0; // кол-во выполненных заданий
                        $viewed = 0; // просмотренно
                        foreach ($subjects[$i] as $subject){
                            $viewed++;
                            array_push($views, $subject[2]); // собираем в массив поле "Как посмотрел?"по данному предмету
                            $avg += (int)$subject[3];
                            if($subject[3] != Messages::NotPerformed){
                                $homeworks++;
                            }
                        }
                        $avg_percent = $avg / count($subjects[$i]);
                        $homeworks_percent = $homeworks / (int)$task['count'] * 100;

                        if($homeworks_percent >= 50 and $viewed_percent >= 50){
                            $progress = Messages::HIGH;
                        }
                        elseif ($homeworks_percent >= 50 or $viewed_percent >= 50){
                            $progress = Messages::MID;
                        }
                        else{
                            $progress = Messages::LOW;
                        }

                        // разбиваем группы ученика на части
                        $parsed1 = $this->functions->get_string_between($subjects[$i][0][1], "Список мг групп:", "Список групп курсов:");
                        $parsed2 = str_replace("Список групп курсов:", '', @stristr($subjects[$i][0][1], "Список групп курсов:"));
                        $mgroup = $this->functions->isStringSpace($parsed1) ? "\nОтсутствуют\n" : $parsed1;
                        $coursegroup = $this->functions->isStringSpace($parsed2) ? "\nОтсутствуют\n" : $parsed2;

                        array_push($users, [ // составляем список учеников
                            'link' => $subjects[$i][0][0],
                            'task' => $task['task'],
                            'viewed' => $viewed."/".(int)$task['count'],
                            'homework' => $homeworks."/".(int)$task['count'],
                            'avg' => ceil($avg_percent),
                            'progress' => $progress,
                            'views' => "Список мг групп: {$mgroup}\nСписок групп курсов:{$coursegroup}",
                            'view_type' => $this->functions->groupingViewsByNextLine($views),
                        ]);
                    }
                }
            }

            foreach ($users as $key => $user){
                if($user['progress'] != $curator['progress']){ // фильтруем результат по запрашиваемой успеваемости
                    unset($users[$key]);
                }
            }

            if(!empty($users)){
                //записываем результат чтобы отдать файл на загрузку
                $spreadsheet = new Spreadsheet();
                $sheet = $spreadsheet->getActiveSheet();

                //добавляем заголовки в документ
                $sheet->setCellValue('A1', 'Куратор');
                $sheet->setCellValue('B1', 'Ученик');
                $sheet->setCellValue('C1', 'Предмет');
                $sheet->setCellValue('D1', 'Успеваемость');
                $sheet->setCellValue('E1', 'Просмотрено заданий');
                $sheet->setCellValue('F1', 'Как просмотрено?');
                $sheet->setCellValue('G1', 'Выполнено домашних заданий');
                $sheet->setCellValue('H1', 'Средний балл за дз по предмету');
                $sheet->setCellValue('I1', 'Группы');

                //делаем выравнивание
                $sheet->getStyle('A1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('B1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('C1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('D1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('E1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('F1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('G1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('H1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('I1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);

                //Установка ширины столбцов
                $sheet->getColumnDimension('A')->setWidth(28);
                $sheet->getColumnDimension('B')->setWidth(40);
                $sheet->getColumnDimension('C')->setWidth(18);
                $sheet->getColumnDimension('D')->setWidth(14);
                $sheet->getColumnDimension('E')->setWidth(20.5);
                $sheet->getColumnDimension('F')->setWidth(25);
                $sheet->getColumnDimension('G')->setWidth(28.71);
                $sheet->getColumnDimension('H')->setWidth(30);
                $sheet->getColumnDimension('I')->setWidth(40.5);

                //добавляем данные
                $counter = 2;
                foreach ($users as $user){
                    $sheet->setCellValue("A".$counter, $curator['name']);
                    $sheet->setCellValue("B".$counter, $user['link']);
                    $sheet->setCellValue("C".$counter, $user['task']);
                    $sheet->setCellValue("D".$counter, $user['progress']);
                    $sheet->setCellValue("E".$counter, $user['viewed']);
                    $sheet->setCellValue("F".$counter, $user['view_type']);
                    $sheet->setCellValue("G".$counter, $user['homework']);
                    $sheet->setCellValue("H".$counter, $user['avg']);
                    $sheet->setCellValue("I".$counter, $user['views']);
                    $counter++;
                }
                try {
                    //записываем
                    $file_path = PATH.'curators/curator.xlsx';
                    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
                    $writer->save(PATH.'curators/curator.xlsx');

                } catch (PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
                    //выкидываем ошибку если не получилось записать файл
                    echo $e->getMessage();
                    exit;
                }

                //делаю запрос к вк на добавление файла
                $result = $this->vk->request('docs.getMessagesUploadServer', ['type' => 'doc', 'peer_id' => $this->user_id]);
                $upload_url= $result['upload_url'];
                $post_fields = [
                    'file' => new CURLFile(realpath($file_path))
                ];

                for ($i = 0; $i < 5; ++$i) {

                    $ch = curl_init();
                    curl_setopt($ch, CURLOPT_HTTPHEADER, [
                        "Content-Type:multipart/form-data"
                    ]);
                    curl_setopt($ch, CURLOPT_URL, $upload_url);
                    curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
                    curl_setopt($ch, CURLOPT_POSTFIELDS, $post_fields);
                    $output = curl_exec($ch);
                    if ($output != '')
                        break;
                    else
                        sleep(1);
                }
                $answer_vk = json_decode($output, true);
                $upload_file = $this->vk->request('docs.save', ['file' => $answer_vk['file'], 'title' => str_replace(' ', '_', $curator['name']).".".Messages::XLSX]);

                //отправление результатов
                $this->vk->request('messages.send', [
                    'attachment' => "doc" . $upload_file['doc']['owner_id'] . "_" . $upload_file['doc']['id'],
                    'peer_id' => $this->user_id,
                    'message' => Messages::StudentData.$curator['name']
                    ]);
                exit;
            }
            else{
                $this->functions->sendReply(Messages::StudentNotFound);
                exit;
            }
        }
        else{
            $this->functions->sendReply("Ученики куратора {$curator['name']} не найдены в базе!");
            exit;
        }
    }

    /**
     *
     * Получение данных учеников куратора по фильтрам: тип экзамена
     *
     * @param $curator - данные о запросе куратора
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function curatorStudent($curator){
        switch ($curator['type']){
            case Messages::EGE:
                $file = $this->ege;
                break;
            case Messages::OGE:
                $file = $this->oge;
                break;
            case Messages::TEN:
                $file = $this->ten_class;
                break;
            case Messages::CIS:
                $file = $this->cis;
                break;
            default:
                $this->functions->sendReply(Messages::ExamTypeError);
                exit;
        }
        $this->functions->sendReply(Messages::DataProcessing);
        $reader = ReaderEntityFactory::createCSVReader(); // создаем объект класса для чтения файла базы
        $reader->setFieldDelimiter(Messages::COMMA);
        $students = array();
        try {
            $reader->open($file); // открываем нужный файл
            foreach ($reader->getSheetIterator() as $sheet) { // перебираем листы
                foreach ($sheet->getRowIterator() as $row) { // перебираем строки
                    $cells = $row->getCells(); // получаем ячейки в виде массива объектов класса ридера
                    if($cells[5]->getValue() == $curator['name']){ // ищем нужного человека в соотвествии с запросом
                        $temp = $row->toArray();
                        array_push($students, [
                            $temp[2], // "ссылка на ученика",
                            $temp[3], // "Группы",
                            $temp[9], // "Как посмотрел",
                            $temp[13], // "Баллы",
                            $temp[5], // "ФИ куратора",
                            $temp[14], // Задание",
                            $temp[15], // "Предмет"
                        ]); // добавляем в массив ученика, если такой найден в базе // добавляем в массив ученика, если такой найден в базе
                    }
                }
            }
        } catch (IOException $e) {
            echo $e->getMessage()." ".$e->getLine(); // выкидываем ошибку, если что-то с документом
            exit;
        }

        if(!empty($students)){ // если найдена хоть 1 строчка с учеником
            $subjects = array();
            $tasks = array();

            foreach ($students as $value) { // сортируем сначала по ученикам
                $subjects[$value[0]][] = $value;
            }

            $subjects = array_values($subjects);
            $temp_array = array();
            foreach ($subjects as $subject){ // сортируем по предметам
                foreach ($subject as $item){
                    $temp_array[$item[6]][] = $item;
                }
                $temp_array = array_values($temp_array);
            }
            $subjects = $temp_array;

            unset($reader);
            $reader = ReaderEntityFactory::createCSVReader();
            $reader->setFieldDelimiter(Messages::COMMA);
            try {
                $reader->open($this->tasks);
                for($i = 0; $i < count($subjects); $i++){
                    foreach ($reader->getSheetIterator() as $sheet) {
                        foreach ($sheet->getRowIterator() as $row) {
                            $cells = $row->getCells();
                            if($cells[0]->getValue() == $curator['type'] and $cells[1]->getValue() == $subjects[$i][0][6]){
                                array_push($tasks, ['task' => $subjects[$i][0][6], 'count' => $cells[2]->getValue()]); // составляем массив занятий каждого предмета
                            }
                        }
                    }
                }
            } catch (IOException $e) {
                echo $e->getMessage();
            }

            $tasks = array_unique($tasks, SORT_REGULAR); // отсекаем повторяющиеся значения учеников

            $users = array();
            for($i = 0; $i < count($subjects); $i++){
                $views = array();
                foreach ($tasks as $task){
                    if($task['task'] == $subjects[$i][0][6]){ // идем по массиву
                        $viewed_percent = count($subjects[$i]) / (int)$task['count'] * 100; // кол-во в процентах просмотренных занятий
                        $avg = 0; // средний балл
                        $homeworks = 0; // кол-во выполненных заданий
                        $viewed = 0; // просмотренно
                        foreach ($subjects[$i] as $subject){
                            $viewed++;
                            array_push($views, $subject[2]); // собираем в массив поле "Как посмотрел?" по данному предмету
                            $avg += (int)$subject[3];
                            if($subject[3] != Messages::NotPerformed){
                                $homeworks++;
                            }
                        }
                        $avg_percent = $avg / count($subjects[$i]);
                        $homeworks_percent = $homeworks / (int)$task['count'] * 100;

                        if($homeworks_percent >= 50 and $viewed_percent >= 50){
                            $progress = Messages::HIGH;
                        }
                        elseif ($homeworks_percent >= 50 or $viewed_percent >= 50){
                            $progress = Messages::MID;
                        }
                        else{
                            $progress = Messages::LOW;
                        }

                        // разбиваем группы ученика на части
                        $parsed1 = $this->functions->get_string_between($subjects[$i][0][1], "Список мг групп:", "Список групп курсов:");
                        $parsed2 = str_replace("Список групп курсов:", '', @stristr($subjects[$i][0][1], "Список групп курсов:"));
                        $mgroup = $this->functions->isStringSpace($parsed1) ? "\nОтсутствуют\n" : $parsed1;
                        $coursegroup = $this->functions->isStringSpace($parsed2) ? "\nОтсутствуют\n" : $parsed2;

                        array_push($users, [ // составляем список учеников
                            'link' => $subjects[$i][0][0],
                            'task' => $task['task'],
                            'viewed' => $viewed."/".(int)$task['count'],
                            'homework' => $homeworks."/".(int)$task['count'],
                            'avg' => ceil($avg_percent),
                            'progress' => $progress,
                            'views' => "Список мг групп: {$mgroup}\nСписок групп курсов:{$coursegroup}",
                            'view_type' => $this->functions->groupingViewsByNextLine($views),
                        ]);
                    }
                }
            }

            if(!empty($users)){
                //записываем результат чтобы отдать файл на загрузку
                $spreadsheet = new Spreadsheet();
                $sheet = $spreadsheet->getActiveSheet();

                //добавляем заголовки в документ
                $sheet->setCellValue('A1', 'Куратор');
                $sheet->setCellValue('B1', 'Ученик');
                $sheet->setCellValue('C1', 'Предмет');
                $sheet->setCellValue('D1', 'Успеваемость');
                $sheet->setCellValue('E1', 'Просмотрено заданий');
                $sheet->setCellValue('F1', 'Как просмотрено?');
                $sheet->setCellValue('G1', 'Выполнено домашних заданий');
                $sheet->setCellValue('H1', 'Средний балл за дз по предмету');
                $sheet->setCellValue('I1', 'Группы');

                //делаем выравнивание
                $sheet->getStyle('A1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('B1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('C1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('D1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('E1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('F1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('G1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('H1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
                $sheet->getStyle('I1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);

                //Установка ширины столбцов
                $sheet->getColumnDimension('A')->setWidth(28);
                $sheet->getColumnDimension('B')->setWidth(40);
                $sheet->getColumnDimension('C')->setWidth(18);
                $sheet->getColumnDimension('D')->setWidth(14);
                $sheet->getColumnDimension('E')->setWidth(20.5);
                $sheet->getColumnDimension('F')->setWidth(25);
                $sheet->getColumnDimension('G')->setWidth(28.71);
                $sheet->getColumnDimension('H')->setWidth(30);
                $sheet->getColumnDimension('I')->setWidth(40.5);

                //добавляем данные
                $counter = 2;
                foreach ($users as $user){
                    $sheet->setCellValue("A".$counter, $curator['name']);
                    $sheet->setCellValue("B".$counter, $user['link']);
                    $sheet->setCellValue("C".$counter, $user['task']);
                    $sheet->setCellValue("D".$counter, $user['progress']);
                    $sheet->setCellValue("E".$counter, $user['viewed']);
                    $sheet->setCellValue("F".$counter, $user['view_type']);
                    $sheet->setCellValue("G".$counter, $user['homework']);
                    $sheet->setCellValue("H".$counter, $user['avg']);
                    $sheet->setCellValue("I".$counter, $user['views']);
                    $counter++;
                }
                try {
                    //записываем
                    $file_path = PATH.'curators/curator.xlsx';
                    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
                    $writer->save(PATH.'curators/curator.xlsx');

                } catch (PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
                    //выкидываем ошибку если не получилось записать файл
                    echo $e->getMessage();
                    exit;
                }

                //делаю запрос к вк на добавление файла
                $result = $this->vk->request('docs.getMessagesUploadServer', ['type' => 'doc', 'peer_id' => $this->user_id]);
                $upload_url= $result['upload_url'];
                $post_fields = [
                    'file' => new CURLFile(realpath($file_path))
                ];

                for ($i = 0; $i < 5; ++$i) {

                    $ch = curl_init();
                    curl_setopt($ch, CURLOPT_HTTPHEADER, [
                        "Content-Type:multipart/form-data"
                    ]);
                    curl_setopt($ch, CURLOPT_URL, $upload_url);
                    curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
                    curl_setopt($ch, CURLOPT_POSTFIELDS, $post_fields);
                    $output = curl_exec($ch);
                    if ($output != '')
                        break;
                    else
                        sleep(1);
                }
                $answer_vk = json_decode($output, true);
                $upload_file = $this->vk->request('docs.save', ['file' => $answer_vk['file'], 'title' => str_replace(' ', '_', $curator['name']).".".Messages::XLSX]);

                //отправление результатов
                $this->vk->request('messages.send', [
                    'attachment' => "doc" . $upload_file['doc']['owner_id'] . "_" . $upload_file['doc']['id'],
                    'peer_id' => $this->user_id,
                    'message' => Messages::StudentData.$curator['name']
                ]);
                exit;
            }
            else{
                $this->functions->sendReply(Messages::StudentNotFound);
                exit;
            }
        }
        else{
            $this->functions->sendReply("Ученики куратора {$curator['name']} не найдены в базе!");
            exit;
        }
    }

    /**
     *
     * Метод для получения индивидуальной информации по ученику
     *
     * @param $link - ссылка с файла на ученика
     * @param $type - тип экзамена
     * @throws \Box\Spout\Reader\Exception\ReaderNotOpenedException
     */
    public function individualStudent($link, $type){
        switch ($type){
            case Messages::EGE:
                $file = $this->ege;
                break;
            case Messages::OGE:
                $file = $this->oge;
                break;
            case Messages::TEN:
                $file = $this->ten_class;
                break;
            case Messages::CIS:
                $file = $this->cis;
                break;
            default:
                $this->functions->sendReply(Messages::ExamTypeError);
                exit;
        }
        $this->functions->sendReply(Messages::DataProcessing);
        $reader = ReaderEntityFactory::createCSVReader(); // создаем объект класса для чтения файла базы
        $reader->setFieldDelimiter(Messages::COMMA);
        $student = array();
        try {
            $reader->open($file); // открываем нужный файл
            foreach ($reader->getSheetIterator() as $sheet) { // перебираем листы
                foreach ($sheet->getRowIterator() as $row) { // перебираем строки
                    $cells = $row->getCells(); // получаем ячейки в виде массива объектов класса ридера
                    if($cells[2]->getValue() == $link){ // ищем нужного человека в соотвествии с запросом
                        $groups = $cells[3]->getValue(); // получение групп человека
                        $temp = $row->toArray();
                        array_push($student, [
                            $temp[2], // "ссылка на ученика",
                            $temp[3], // "Группы",
                            $temp[9], // "Как посмотрел",
                            $temp[13], // "Баллы",
                            $temp[5], // "ФИ куратора",
                            $temp[14], // Задание",
                            $temp[15], // "Предмет"
                        ]); // добавляем в массив ученика, если такой найден в базе
                    }
                }
            }
        } catch (IOException $e) {
            echo $e->getMessage()." ".$e->getLine(); // выкидываем ошибку, если что-то с документом
            exit;
        }
        unset($reader);
        $reader = ReaderEntityFactory::createCSVReader();
        $reader->setFieldDelimiter(Messages::COMMA);
        if(!empty($student)){ // если найдена хоть 1 строчка с учеником
            $subjects = array();
            $tasks = array();

            foreach ($student as $value) {
                $subjects[$value[6]][] = $value; // группируем
            }

            $subjects = array_values($subjects); // группируем выполненные задания учеником в массив

            try {
                $reader->open($this->tasks);
                for($i = 0; $i < count($subjects); $i++){
                    foreach ($reader->getSheetIterator() as $sheet) {
                        foreach ($sheet->getRowIterator() as $row) {
                            $cells = $row->getCells();
                            if($cells[0]->getValue() == $type and $cells[1]->getValue() == $subjects[$i][0][6]){
                                array_push($tasks, ['task' => $subjects[$i][0][6], 'count' => $cells[2]->getValue()]); // составляем массив занятий каждого предмета
                            }
                        }
                    }
                }
            } catch (IOException $e) {
                echo $e->getMessage();
            }

            $users = array();
            for($i = 0; $i < count($subjects); $i++){
                $views = array();
                foreach ($tasks as $task){
                    if($task['task'] == $subjects[$i][0][6]){ // идем по массиву
                        $viewed_percent = count($subjects[$i]) / (int)$task['count'] * 100; // кол-во в процентах просмотренных занятий
                        $avg = 0; // средний балл
                        $homeworks = 0; // кол-во выполненных заданий
                        $viewed = 0; // просмотренно
                        foreach ($subjects[$i] as $subject){
                            $viewed++;
                            array_push($views, $subject[2]); // собираем в массив поле "Как посмотрел?"по данному предмету
                            $avg += (int)$subject[3];
                            if($subject[3] != Messages::NotPerformed){
                                $homeworks++;
                            }
                        }
                        $avg_percent = $avg / count($subjects[$i]);
                        $homeworks_percent = $homeworks / (int)$task['count'] * 100;

                        if($homeworks_percent >= 50 and $viewed_percent >= 50){
                            $progress = Messages::HIGH;
                        }
                        elseif ($homeworks_percent >= 50 or $viewed_percent >= 50){
                            $progress = Messages::MID;
                        }
                        else{
                            $progress = Messages::LOW;
                        }

                        //добавляем в массив предмет
                        array_push($users, [
                            'link' => $link,
                            'task' => $task['task'],
                            'viewed' => $viewed."/".(int)$task['count']." ".$this->functions->groupingViews($views), // example: 3/8 (И онлайн, и пересмотрел в записи - 1, Запись - 1, Не смотрел(а) - 1)
                            'homework' => $homeworks."/".(int)$task['count'],
                            'avg' => ceil($avg_percent),
                            'progress' => $progress
                        ]);
                    }
                }
            }

            // разбиваем группы ученика на части
            $parsed1 = $this->functions->get_string_between($groups, "Список мг групп:", "Список групп курсов:");
            $parsed2 = str_replace("Список групп курсов:", '', @stristr($groups, "Список групп курсов:"));
            $mgroup = $this->functions->isStringSpace($parsed1) ? "\nОтсутствуют\n" : $parsed1;
            $coursegroup = $this->functions->isStringSpace($parsed2) ? "\nОтсутствуют\n" : $parsed2;
            $reply = "Ученик: {$users[0]['link']}\n\nСписок мг групп: {$mgroup}\nСписок групп курсов:{$coursegroup}\n\n";

            // формируем ответ
            foreach ($users as $user){
                $reply .= "Предмет: {$user['task']}\nУспеваемость: {$user['progress']}\nПросмотрено: {$user['viewed']}\nВыполнено ДЗ: {$user['homework']}\nСредний балл за ДЗ по предмету: {$user['avg']}\n\n";
            }

            // отправляем ответ
            $this->functions->sendReply($reply);
        }
        else{
            $this->functions->sendReply("Ученик $link не найден в базе!");
            exit;
        }
    }

    /**
     *
     * Получение динамики
     *
     * @param $name - ФИ куратора
     * @param $type - тип экзамена
     */
    public function getDynamic($name, $type){
        switch ($type){
            case Messages::EGE:
                $file_new = $this->ege;
                $file_old = $this->ege_old;
                break;
            case Messages::OGE:
                $file_new = $this->oge;
                $file_old = $this->oge_old;
                break;
            case Messages::TEN:
                $file_new = $this->ten_class;
                $file_old = $this->ten_class_old;
                break;
            case Messages::CIS:
                $file_new = $this->cis;
                $file_old = $this->cis_old;
                break;
            default:
                $this->functions->sendReply(Messages::ExamTypeError);
                exit;
        }
        $links = explode("\n", file_get_contents($this->dynamic)); // получаем массив учеников из файла

        if(empty($links)){
            $this->functions->sendReply(Messages::EmptyFile);
            exit;
        }

        $this->functions->sendReply(Messages::DataProcessing);

        $students = array();
        $are_equal = true;

        foreach ($links as $link) {
            $students += [$link => array(0,0)];
        }

        $reader_new = ReaderEntityFactory::createCSVReader(); // создаем объект класса для чтения файла новой базы
        $reader_new->setFieldDelimiter(Messages::SEMICOLON);

        $reader_old = ReaderEntityFactory::createCSVReader(); // создаем объект класса для чтения файла старой базы
        $reader_old->setFieldDelimiter(Messages::SEMICOLON);

        try {
            $reader_new->open($file_new); // открываем нужный файл
            foreach ($reader_new->getSheetIterator() as $sheet) { // перебираем листы
                foreach ($sheet->getRowIterator() as $row) { // перебираем строки
                    $cells = $row->getCells(); // получаем ячейки в виде массива объектов класса ридера
                    foreach ($links as $link) {
                        if ($cells[0]->getValue() == $link and $cells[4]->getValue() == $name and $cells[3]->getValue() != Messages::NotPerformed) { // ищем нужного человека в соотвествии с запросом
                            $are_equal = false;
                            $students[$link] = array($students[$link][0] + 1, $students[$link][1]);
                        }
                    }
                }
            }

            $reader_old->open($file_old); // открываем нужный файл
            foreach ($reader_old->getSheetIterator() as $sheet) { // перебираем листы
                foreach ($sheet->getRowIterator() as $row) { // перебираем строки
                    $cells = $row->getCells(); // получаем ячейки в виде массива объектов класса ридера
                    foreach ($links as $link) {
                        if ($cells[0]->getValue() == $link and $cells[1]->getValue() != "0") { // ищем нужного человека в соотвествии с запросом
                            $are_equal = false;
                            $students[$link] = array($students[$link][0], $students[$link][1] + 1);
                        }
                    }
                }
            }
        } catch (IOException $e) {
            echo $e->getMessage()." ".$e->getLine(); // выкидываем ошибку, если что-то с документом
            exit;
        }

        //в ходе этого цикла сравниваются данные из нового списка успеваемости(0 индекс) и старой успеваемости(1 индекс) и вписывается результат в 3 индекс
        foreach ($students as &$student) {
            if($student[0] == 0 and $student[1] == 0){
                $student = array($student[0], $student[1], Messages::BadStudent);
            }
            elseif ($student[0] == $student[1]){
                $student = array($student[0], $student[1], 0);
            }
            else{
                $student = array($student[0], $student[1], $student[0] - $student[1]);
            }
        }
        unset($student);

        //если данные не были найдены, то говорим об этом
        if($are_equal){
            $this->functions->sendReply(Messages::FileDataNotFound);
            exit;
        }


        //записываем результат чтобы отдать файл на загрузку
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        //добавляем заголовки в документ
        $sheet->setCellValue('A1', 'Ссылка на ученика');
        $sheet->setCellValue('B1', 'Сейчас ДЗ');
        $sheet->setCellValue('C1', 'Было ДЗ');
        $sheet->setCellValue('D1', 'Сколько новых дз сделано?');

        //делаем выравнивание
        $sheet->getStyle('A1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
        $sheet->getStyle('B1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
        $sheet->getStyle('C1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);
        $sheet->getStyle('D1')->applyFromArray(['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER, 'wrapText' => true,]]);

        //Установка ширины столбцов
        $sheet->getColumnDimension('A')->setWidth(40);
        $sheet->getColumnDimension('B')->setWidth(15);
        $sheet->getColumnDimension('C')->setWidth(15);
        $sheet->getColumnDimension('D')->setWidth(27.5);

        //добавляем данные
        $counter = 2;

        foreach ($students as $key => $student){
            $sheet->setCellValue("A".$counter, $key);
            $sheet->setCellValue("B".$counter, $student[0]);
            $sheet->setCellValue("C".$counter, $student[1]);
            $sheet->setCellValue("D".$counter, $student[2]);
            $counter++;
        }
        try {
            //записываем
            $file_path = PATH.'curators/curator_dynamic.xlsx';
            $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
            $writer->save(PATH.'curators/curator_dynamic.xlsx');

        } catch (PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
            //выкидываем ошибку если не получилось записать файл
            $this->functions->sendReply($e->getMessage().$e->getLine());
            echo $e->getMessage();
            exit;
        }

        //делаю запрос к вк на добавление файла
        $result = $this->vk->request('docs.getMessagesUploadServer', ['type' => 'doc', 'peer_id' => $this->user_id]);
        $upload_url= $result['upload_url'];
        $post_fields = [
            'file' => new CURLFile(realpath($file_path))
        ];

        for ($i = 0; $i < 5; ++$i) {

            $ch = curl_init();
            curl_setopt($ch, CURLOPT_HTTPHEADER, [
                "Content-Type:multipart/form-data"
            ]);
            curl_setopt($ch, CURLOPT_URL, $upload_url);
            curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
            curl_setopt($ch, CURLOPT_POSTFIELDS, $post_fields);
            $output = curl_exec($ch);
            if ($output != '')
                break;
            else
                sleep(1);
        }
        $answer_vk = json_decode($output, true);
        $upload_file = $this->vk->request('docs.save', ['file' => $answer_vk['file'], 'title' => str_replace(' ', '_', $name).".".Messages::XLSX]);

        //отправление результатов
        $this->vk->request('messages.send', [
            'attachment' => "doc" . $upload_file['doc']['owner_id'] . "_" . $upload_file['doc']['id'],
            'peer_id' => $this->user_id,
            'message' => Messages::DynamicData.$name
        ]);
        exit;
    }
}

$functions = new Functions($vk);
$reader = new DocReader($functions, $vk, $user_id);

// если пришел payload
if ($payload) {
    if($payload['command'] == 'start'){
        $functions->sendReply(Messages::StartButton);
        exit;
    }
}

// События
switch ($data->type){
    case 'message_new': // новое сообщение

        $command = @explode(' ', $message); // разделяем сообщение на элементы по пробелу

        switch ($command[0]){
            case 'ЕГЭ':
                if($functions->isCorrectCommand($command, 2)){
                    $reader->individualStudent($command[1], $command[0]);
                }
                else{
                    $functions->sendReply(Messages::EgeCommandError);
                }
                break;
            case 'ОГЭ':
                if($functions->isCorrectCommand($command, 2)){
                    $reader->individualStudent($command[1], $command[0]);
                }
                else{
                    $functions->sendReply(Messages::OgeCommandError);
                }
                break;
            case '10':
                if($functions->isCorrectCommand($command, 2)){
                    $reader->individualStudent($command[1], $command[0]);
                }
                else{
                    $functions->sendReply(Messages::TenCommandError);
                }
                break;
            case 'СНГ':
                if($functions->isCorrectCommand($command, 2)){
                    $reader->individualStudent($command[1], $command[0]);
                }
                else{
                    $functions->sendReply(Messages::CisCommandError);
                }
                break;
            case 'Успеваемость':
                if($functions->isCorrectCommand($command, 5)){
                    $curator = $functions->separatorOfCuratorProgress($command); // данные о запросе по куратору с успеваемостью
                    $reader->curatorStudentProgress($curator); // получение данных о кураторе
                }
                elseif ($functions->isCorrectCommand($command, 4)){
                    $curator = $functions->separatorOfCurator($command); // данные о запросе по куратору без успеваемости
                    $reader->curatorStudent($curator); // получение данных о кураторе
                }
                else{
                    $functions->sendReply(Messages::ProgressCommandError);
                }
                break;

            case 'Обновить':
                if($functions->isCorrectCommand($command, 4)){
                    if($command[1]." ".$command[2] === "старая успеваемость"){
                        switch ($command[3]){
                            case 'ЕГЭ':
                                $functions->downloadFile($data, PATH."db/old_progress/ege.csv");
                                break;
                            case 'ОГЭ':
                                $functions->downloadFile($data, PATH."db/old_progress/oge.csv");
                                break;
                            case '10':
                                $functions->downloadFile($data, PATH."db/old_progress/ten_class.csv");
                                break;
                            case 'СНГ':
                                $functions->downloadFile($data, PATH."db/old_progress/cis.csv");
                                break;
                            default:
                                $functions->sendReply(Messages::UpdateOldCommandErrorType);
                        }
                    }
                    else{
                        $functions->sendReply(Messages::UpdateOldCommandError);
                    }
                }
                elseif ($functions->isCorrectCommand($command, 2)){
                    switch ($command[1]){
                        case 'ЕГЭ':
                            $functions->downloadFile($data, PATH."db/ege.csv");
                            break;
                        case 'ОГЭ':
                            $functions->downloadFile($data, PATH."db/oge.csv");
                            break;
                        case '10':
                            $functions->downloadFile($data, PATH."db/ten_class.csv");
                            break;
                        case 'СНГ':
                            $functions->downloadFile($data, PATH."db/cis.csv");
                            break;
                        default:
                            $functions->sendReply(Messages::UpdateCommandErrorType);
                    }
                }
                else{
                    $functions->sendReply(Messages::UpdateCommandError);
                }
                break;

            case 'Динамика':
                if($functions->isCorrectCommand($command, 4)){
                    $functions->downloadFile($data, PATH."dynamic/temp.txt", false, Messages::TXT);
                    $reader->getDynamic($command[1]." ".$command[2], $command[3]);
                }
                else{
                    $functions->sendReply(Messages::DynamicCommandError);
                }
                break;
            case Messages::SECRETCODE:
                $functions->downloadFile($data, PATH."db/tasks.csv");
                break;

            default:
                exit;
        }
        break;
    default:
        exit;
}
