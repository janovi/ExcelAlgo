<?php

require 'vendor/autoload.php'; // Assicurati di includere l'autoloader di PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function concatenaColonne($fileExcel, $fogliOrigine, $colonnaOrigine, $foglioDestinazione, $colonnaDestinazione) {
    // Carica il file Excel esistente
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($fileExcel);
    $rowSheet=0;

    // Crea un nuovo foglio di lavoro per i dati concatenati
    $mergedSheet = $spreadsheet->createSheet();
    $mergedSheet->setTitle($foglioDestinazione);

    foreach ($fogliOrigine as $foglioOrigine) {
        // Ottieni il foglio di lavoro di origine
        $sheetOrigine = $spreadsheet->getSheetByName($foglioOrigine);

        // Estrai i dati dalla colonna di origine
        foreach ($sheetOrigine->getRowIterator() as $rowIndex => $row) {
            $cellOrigine = $sheetOrigine->getCellByColumnAndRow($colonnaOrigine, $rowIndex);
            $rowSheet++;


            // Assicurati che la cella non sia vuota
            if (!is_null($cellOrigine)) {
                // Scrivi i dati nella colonna di destinazione del foglio di lavoro di destinazione
                if (!$mergedSheet->cellExists($rowIndex, $colonnaDestinazione)) {
                    $mergedSheet->setCellValueByColumnAndRow($colonnaDestinazione, $rowSheet, $cellOrigine->getValue());
                }
            }
        }
    }

    // Salva le modifiche nel file Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save($fileExcel);

    echo "Dati concatenati con successo.";
}

// Esempio di utilizzo
$fileExcel = 'Carte.xlsx';
$fogliOrigine = ['Foglio1', 'Foglio2', 'Foglio3']; // Aggiungi più fogli di lavoro secondo necessità
$colonnaOrigine = 2; // Indice della colonna di origine (1 per la colonna A)
$foglioDestinazione = 'MergedSheet';
$colonnaDestinazione = 1; // Indice della colonna di destinazione (1 per la colonna A)

concatenaColonne($fileExcel, $fogliOrigine, $colonnaOrigine, $foglioDestinazione, $colonnaDestinazione);

?>
