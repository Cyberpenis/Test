<?php 


require 'C:\Ampps\php\vendor\autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
  $phpWord = new PhpWord();


  $tableStyle = [
      'borderSize' => 6,
      'borderColor' => '000000',
  ];
  
  $section = $phpWord->addSection();
  $table = [
    [
        ['value' => 'Коэффициенты платежеспособности и ликвидности', 'colspan' => 1, 'rowspan' => 2],
        ['value' => 'Значения', 'colspan' => 3, 'rowspan' => 1],
        ['value' => 'Отклонение', 'colspan' => 3, 'rowspan' => 1],
    ],
    [
        ['value' => '31.12.2023', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '31.12.2022', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '31.12.2021', 'colspan' => 1, 'rowspan' => 1],
        ['value' => 'На конец 2023 г.\n(базисные)', 'colspan' => 1, 'rowspan' => 1],
        ['value' => 'На конец 2023 г.\n(цепные)', 'colspan' => 1, 'rowspan' => 1],
        ['value' => 'На конец 2022 г.', 'colspan' => 1, 'rowspan' => 1],
    ],
    [
        ['value' => 'Общий показатель ликвидности', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
    ],
    [
        ['value' => 'Коэффициент абсолютной ликвидности', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
    ],
    [
        ['value' => 'Коэффициент "критической" ликвидности', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
    ],
    [
        ['value' => 'Коэффициент текущей ликвидности', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
    ],
    [
        ['value' => 'Коэффициент маневренности функционирующего капитала', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
    ],
    [
        ['value' => 'Доля оборотных средств в активах', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
    ],
    [
        ['value' => 'Коэффициент обеспеченности собственными средствами', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
        ['value' => '0.00', 'colspan' => 1, 'rowspan' => 1],
    ],
];

  $phpTable = $section->addTable($tableStyle);
  
  $rowCounter = 0;
  $prevMerged = false;

  foreach ($table as $row) {
      $tableRow = $phpTable->addRow();
      $colCounter = 0;
      $colMaxCount = count($row);

      $upperRowCheck = 0;

      while ($colCounter < $colMaxCount) {
          if ($rowCounter === 0) {
              if ($table[$rowCounter][$colCounter]['rowspan'] > 1) {
                  $tableRow->addCell(2000, [
                      'borderSize' => 6,
                      'borderColor' => '000000',
                      'vMerge' => 'restart',
                      'gridSpan'=> $table[$rowCounter][$colCounter]['colspan']
                  ])->addText($table[$rowCounter][$colCounter]['value']);
                  $prevMerged = true;
              } else {
                  $tableRow->addCell(2000, [
                      'borderSize' => 6,
                      'borderColor' => '000000',
                      'gridSpan'=> $table[$rowCounter][$colCounter]['colspan']
                  ])->addText($table[$rowCounter][$colCounter]['value']);
              }
              $colCounter++;
          } elseif ($rowCounter === 1 && $prevMerged){
                if ($table[$rowCounter-1][$upperRowCheck]['rowspan']>1){

                    $tableRow->addCell(2000, [
                        'borderSize' => 6,
                        'borderColor' => '000000',
                        'vMerge' => 'continue',
                        'gridSpan'=> $table[$rowCounter][$upperRowCheck]['colspan']
                    ]);
                    
                } else {
                    if ($table[$rowCounter-1][$upperRowCheck]['colspan'] > 1){
                        for ($n=0; $n <$table[$rowCounter-1][$upperRowCheck]['colspan']; $n++){
                            $tableRow->addCell(2000, [
                                'borderSize' => 6,
                                'borderColor' => '000000',
                                'gridSpan'=> $table[$rowCounter][$colCounter]['colspan']
                            ])->addText(str_replace('\n', '</w:t><w:br/><w:t>', $table[$rowCounter][$colCounter]['value']));
                            $colCounter+= $table[$rowCounter][$colCounter]['colspan'];
                        }
                    } else {
                    $tableRow->addCell(2000, [
                        'borderSize' => 6,
                        'borderColor' => '000000',
                        'gridSpan'=> $table[$rowCounter][$colCounter]['colspan']
                    ])->addText($table[$rowCounter][$colCounter]['value']);
                    }
                }
            $upperRowCheck++;
            } else {
                $tableRow->addCell(2000, [
                  'borderSize' => 6,
                  'borderColor' => '000000',
                  'gridSpan'=> $table[$rowCounter][$colCounter]['colspan']
            ])->addText($table[$rowCounter][$colCounter]['value']);
            $colCounter++;
            }
          }
          $rowCounter++;
      }


  $section->addTextBreak();


    $fileName = 'tables.docx';
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment; filename="' . $fileName . '"');

    $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save('php://output');

}

?>
