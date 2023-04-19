<?php

/*Libs relacionadas*/
use PhpOffice\PhpPresentation\Style\Alignment; //classe de estilo de alinhamentos
use PhpOffice\PhpPresentation\Style\Color; //classe de estilo de cores

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
}