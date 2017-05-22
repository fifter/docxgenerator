<?php

include "docxDocument.php";

function deleteNode($node) {
    deleteChildren($node);
    $parent = $node->parentNode;
    $parent->removeChild($node);
}

function deleteChildren($node) {
    while (isset($node->firstChild)) {
        deleteChildren($node->firstChild);
        $node->removeChild($node->firstChild);
    }
}

interface DocumentGeneratorInterface {

    /**
     * @param string $filename Имя файла шаблона для загрузки и использования
     */
    public function setTemplate($filename);

    /**
     * @param mixed $data Произвольные данные, которые нужно уложить в документ
     * Пример:
     * [
     *  'company' => 'ООО "Счастье"',
     *  'sum' => '100',
     *  'items' => [
     *      [
     *          'name' => 'Хлеб',
     *          'sum' => 20,
     *      ],
     *      [
     *          'name' => 'Колбаса',
     *          'sum' => 80,
     *      ],
     *  ],
     * ]
     */
    public function setData($data);

    /**
     * @param string $filename Имя/путь для сохранения сгенерированного файла
     */
    public function saveToFile($filename);

}


class DocxGenerator implements DocumentGeneratorInterface{

    private $file, $data;


    private static function changeContent($file, $data)
    {
        if($file instanceof DocxDocument)
            $content = $file->getContent();
        else
            throw new Exception('Параметр функции не является объектом docx.');


        $dom = new DOMDocument();
        $dom->loadXML($content);

        $elements = $dom->getElementsByTagName("*");
        $state = "out";
        $startTag = $endTag = 0;

        $startCharPosition = $endCharPosition = 0;
        $tags = [];
        $varName = "";
        for($i = 0; $i < $elements->length; $i++)
        {
            $el = $elements->item($i);
            if($el->nodeName === "w:t")
            {
                for($k = 0; $k < strlen($el->nodeValue); $k++)
                {
                    switch($el->nodeValue[$k])
                    {
                        case "{": {
                            $startTag = $i;
                            $startCharPosition = $k;
                            $state = "text";
                            $varName = "";
                            break;
                        }
                        case "}": {
                            $state = "out";
                            $endTag = $i;
                            $endCharPosition = $k;

                            if (isset($data[$varName])) {
                                if ($startTag === $endTag) {
                                    $string = $elements->item($startTag)->nodeValue;
                                    $elements->item($startTag)->nodeValue =
                                        substr($string, 0, $startCharPosition)
                                        . $data[$varName] . substr($string, $endCharPosition + 1);
                                } else {
                                    $startNode = $elements->item($startTag);
                                    $endNode = $elements->item($endTag);
                                    $otherNodes = [];
                                    foreach ($tags as $tag)
                                        $otherNodes[] = $elements->item($tag);

                                    $startNode->nodeValue =
                                        substr($elements->item($startTag)->nodeValue, 0, $startCharPosition);
                                    if ($startNode->nodeValue === "")
                                        deleteNode($startNode->parentNode);

                                    if (isset($tags[0])) {
                                        $endNode->nodeValue =
                                            substr($endNode->nodeValue, $endCharPosition + 1);
                                        if ($endNode->nodeValue === "")
                                            deleteNode($endNode->parentNode);

                                        array_shift($otherNodes)->nodeValue = $data[$varName];
                                        foreach ($otherNodes as $node) {
                                            deleteNode($node->parentNode);
                                        }
                                    } else {
                                        $endNode->nodeValue =
                                            $data[$varName] .
                                            substr($endNode->nodeValue, $endCharPosition + 1);
                                    }
                                }
                            }
                            $tags = [];
                            $varName = "";
                            break;
                        }
                        default:{
                            switch($state)
                            {
                                case "text":
                                    if($startTag !== $i)
                                    {
                                        if($tags === [])
                                            $tags[]= $i;
                                        else
                                            foreach($tags as $tag)
                                                if($tag !== $i)
                                                $tags[] = $i;

                                    }
                                    $varName .= $el->nodeValue[$k];
                                    break;
                                case "out":
                                    break;
                            }
                            break;
                        }
                    }
                }
            }
        }

        $file->setContent($dom->saveXML());
    }

    
    public function setTemplate($filename)
    {
        if(file_exists($filename))
        {
            $file = new DocxDocument($filename);
            if(!$file->hasContent())
            {
                throw new Exception('Отсутствует контент в документе.');
            }
            $this->file = $file;
        }
        else
            throw new Exception('Файл не существует.');
    }


    public function setData($data)
    {
        $this->data = $data;
    }


    public function saveToFile($filename)
    {
        if(!isset($this->file))
        {
            throw new Exception('Не выбран файл шаблона.');
        }
        if(!isset($this->data))
        {
            throw new Exception('Нет данных для заполнения шаблона.');
        }

        $newFile = new DocxDocument($filename);
        $newFile->copyDoc($this->file);

        DocxGenerator::changeContent($newFile, $this->data);
    }


    /**
     * @param string $template Имя файла шаблона для загрузки и использования
     *
     * @param mixed $data Произвольные данные, которые нужно уложить в документ
     * Пример:
     * [
     *  'company' => 'ООО "Счастье"',
     *  'sum' => '100',
     *  'items' => [
     *      [
     *          'name' => 'Хлеб',
     *          'sum' => 20,
     *      ],
     *      [
     *          'name' => 'Колбаса',
     *          'sum' => 80,
     *      ],
     *  ],
     * ]
     *
     * @param string $saveTo Имя/путь для сохранения сгенерированного файла
     */
    public static function generateDocument($template, $data, $saveTo)
    {
        $file = new DocxDocument($template);
        $newFile = new DocxDocument($saveTo);
        $newFile->copyDoc($file);

        DocxGenerator::changeContent($newFile, $data);

        $file->close();
        $newFile->close();
    }
}
