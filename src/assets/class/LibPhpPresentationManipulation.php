<?php

/*Libs relacionadas*/
use PhpOffice\PhpPresentation\Style\Alignment; //classe de estilo de alinhamentos
use PhpOffice\PhpPresentation\Style\Color; //classe de estilo de cores
use PhpOffice\PhpPresentation\PhpPresentation; //classe do PhpPresentation
use PhpOffice\PhpPresentation\Slide\Background\Image; //utilizacao de imagens
use PhpOffice\PhpPresentation\IOFactory; //classe para manipular os arquivos

class LibPhpPresentationManipulation
{

    //desc: criacao de texto
    //params:
    //return: (obj) RichTextShape
    public function create_text($slide, $height, $width, $offsetX, $offsetY, $text, $isBold, $fontSize, $color){
        //texto
        $title = $slide->createRichTextShape() //cria forma (texto)
        ->setHeight($height) //altura
        ->setWidth($width) //largura
        ->setOffsetX($offsetX) //posicao em relacao ao eixo X
        ->setOffsetY($offsetY); //posicao em relacao ao eixo Y
        //alinhamento do paragrafo (horizontal)
        $title->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        //texto do paragrafo
        $text = $title->createTextRun($text);
        //fonte do texto como negrito
        $text->getFont()->setBold($isBold)
            ->setSize($fontSize) //tamanho da fonte
            ->setColor(new Color($color)); //cor da fonte
    }

    //desc: cria apresentacao
    //params: nenehum
    //return: (obj) PhpPresentation
    static function new_presentation(){
        return new PhpPresentation();
    }

    //desc: cria slide
    //params: (obj) PhpPresentation
    //return: (obj) getActiveSlide
    static function new_slide($presentation){
        return $presentation->getActiveSlide();
    }

    //desc: cria imagem
    //params: (obj) getActiveSlide, (string) caminho de imagem
    //return: nenhum
    function set_background_image_in_slide($slide, $file_name){
        $bg_image = new Image();
        $bg_image->setPath(IMAGE_STORAGE."/$file_name");
        $slide->setBackground($bg_image);
    }

    //desc: cria arquivo
    //params: (obj) PhpPresentation, (string) tipo de arquivo, (string) nome do arquivo
    //return
    function create_pptx_file($presentation, $file_type, $file_name){
        $oWriterPPTX = IOFactory::createWriter($presentation, $file_type); //definindo o tipo de arquivo como PowerPoint2007
        $oWriterPPTX->save(PRESENTATION_STORAGE. "/$file_name");
    }
}