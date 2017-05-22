<?php

include "documentGenerator.php";


try {
    
    DocxGenerator::generateDocument("example.docx",
        ["name" => "ilili",
            "id" => "123",
            "kipa" => "rt67",
            "kipaqwe" => ['a', 't'],
        ]
        , 'asa.docx');


    $invoice = new DocxGenerator();
    $invoice->setTemplate('sampleInvoice.docx');
    $invoice->setData([
        'no' => 14,
        'date' => '28.06.2016',
        'until' => '20.06.2016',
        'subject' => 'HelpOnClick app',
        'sum' => '5067.02',
    ]);
    $invoice->saveToFile('invoice.docx');


 }
catch (Exception $e) {
    echo $e;
}