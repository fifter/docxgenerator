<?php

class DocxDocument{



    // Содержимое документа
    
    private $zipData;

    public function __construct($filename){

        $this->zipData = new ZipArchive();

        if(file_exists($filename))
        {
            if ( $this->zipData->open($filename) !== TRUE)
            {
                die("Unable to open {$filename}\n");
            }
        }
        else
        {
            if ( $this->zipData->open($filename, ZIPARCHIVE::CREATE) !== TRUE)
            {
                die("Unable to open {$filename}\n");
            }
        }
    }

    public function copyDoc($doc)
    {
        if($doc instanceof DocxDocument)
        {
            for($i=0; $i < $doc->zipData->numFiles; $i++)
            {
                $this->zipData->addFromString($doc->zipData->getNameIndex($i),
                   $doc->zipData->getFromIndex($i));
            }
            $name = $this->zipData->filename;
            $this->zipData->close();
            $this->zipData->open($name);
        }
        else
            throw new Exception('Параметр функции не является объектом docx.');
    }

    public function hasContent(){
        if($this->zipData->getFromName("word/document.xml") !== FALSE)
            return TRUE;
        else
            return FALSE;
        
    }
    public function getContent()
    {
        return $this->zipData->getFromName("word/document.xml");
    }

    public function setContent($content)
    {
        $this->zipData->deleteName("word/document.xml");
        $this->zipData->addFromString("word/document.xml", $content);
    }
    public function close()
    {
        return $this->zipData->close();
    }

}
