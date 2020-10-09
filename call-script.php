<?php

require_once 'XlsExchange.php';

(new XlsExchange())
        ->setInputFile('order.json')
        ->setOutputFile('items.xlsx')
        ->export();