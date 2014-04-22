<?php

namespace Ephp\OfficeBundle;

use Symfony\Component\HttpFoundation\StreamedResponse;

/**
 * Factory for PHPWord objects, StreamedResponse, and PHPExcel_Writer_IWriter.
 *
 * @package Ephp\OfficeBundle
 */
class FactoryWord
{
    private $phpWordIO;

    public function __construct($phpWordIO = '\PhpOffice\PhpWord\IOFactory')
    {
        $this->phpWordIO = $phpWordIO;
    }
    /**
     * Creates an empty PHPExcel Object if the filename is empty, otherwise loads the file into the object.
     *
     * @param string $filename
     *
     * @return \PHPExcel
     */
    public function createPHPWordObject($filename =  null)
    {
        if (null == $filename) {
            $phpWordObject = new \PhpOffice\PhpWord\PhpWord();

            return $phpWordObject;
        }

        return call_user_func(array($this->phpWordIO, 'load'), $filename);
    }

    /**
     * Create a writer given the PHPExcelObject and the type,
     *   the type coul be one of PHPExcel_IOFactory::$_autoResolveClasses
     *
     * @param \PHPWord $phpWordObject
     * @param string    $type
     *
     *
     * @return \PHPWord_Writer_IWriter
     */
    public function createWriter(\PhpOffice\PhpWord\PhpWord $phpWordObject, $type = 'Word2007')
    {
        return call_user_func(array($this->phpWordIO, 'createWriter'), $phpWordObject, $type);
    }

    /**
     * Stream the file as Response.
     *
     * @param \PHPWord_Writer_IWriter $writer
     * @param int                      $status
     * @param array                    $headers
     *
     * @return StreamedResponse
     */
    public function createStreamedResponse(\PhpOffice\PhpWord\Writer\IWriter $writer, $status = 200, $headers = array())
    {
        return new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            },
            $status,
            $headers
        );
    }
}
