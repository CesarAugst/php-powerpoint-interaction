<?php

/*Libs relacionadas*/
use PhpOffice\PhpPresentation\Style\Alignment; //classe de estilo de alinhamentos
use PhpOffice\PhpPresentation\Style\Color; //classe de estilo de cores
use PhpOffice\PhpPresentation\PhpPresentation; //classe do PhpPresentation
use PhpOffice\PhpPresentation\Slide\Background\Image; //utilizacao de imagens
use PhpOffice\PhpPresentation\IOFactory; //classe para manipular os arquivos
use PhpOffice\PhpPresentation\Style\Bullet; //Bullet



class LibPhpPresentationManipulation
{
    //desc: cracao de box para o texto
    //params: (string) type box, (obj) getActiveSlide, (number) altura, (number) largura, (number) posicao eixo X, (number) posicao eixo y
    //return: (obj) createRichTextShape
    static function create_box($created_box, $height, $width, $offsetX, $offsetY){
        //espaco ocupado pela forma
        $created_box->setHeight($height); //altura
        $created_box->setWidth($width); //largura
        $created_box->setOffsetX($offsetX); //posicao em relacao ao eixo X
        $created_box->setOffsetY($offsetY); //posicao em relacao ao eixo Y
        //retorna a box apos formacao
        return $created_box;
    }

    //desc: criacao de texto
    //params: (obj) getActiveSlide, (obj) alinhamento, (string) texto, (bool) se bold, (number) fonte0size, (string) color
    //return: (obj) RichTextShape
    static public function create_text($created_box, $alignment, $text, $isBold, $fontSize, $color){
        //cria box de texto
        $shape = $created_box;
        //alinhamento
        $shape->getActiveParagraph()->getAlignment()->setHorizontal($alignment);
        //texto do paragrafo
        $textRun = $shape->createTextRun($text);
        //fonte do texto como negrito
        $textRun->getFont()->setBold($isBold)
            ->setSize($fontSize) //tamanho da fonte
            ->setColor(new Color($color)); //cor da fonte
        //retorna o texto
        return $textRun;
    }

    //desc: criacao de texto com paragrafos
    //params: (obj) getActiveSlide, (obj) alinhamento, (array<string>) texto, (bool) se bold, (number) fonte0size, (string) color, (obj) bullet
    //return: (obj) RichTextShape
    static public function create_paragraph_text($created_box, $alignment_type, $array_text, $isBold, $fontSize, $color, $bullet_type){
        //cria box de texto
        $shape = $created_box;
        //alinhamento
        $shape->getActiveParagraph()->getAlignment()->setHorizontal($alignment_type);
        //fonte do texto como negrito
        $shape->getActiveParagraph()->getFont()->setBold($isBold) //negrito
            ->setSize($fontSize) //tamanho da fonte
            ->setColor(new Color($color)); //cor da fonte
        $shape->getActiveParagraph()->getBulletStyle()->setBulletType($bullet_type);
        //texto
        foreach ($array_text as $index => $text){
            if($index == 0){
                $shape->createTextRun($text);
            }else{
                $shape->createParagraph()->createTextRun($text);
            }
        }
        //retorna o paragrafo
        return $shape;
    }

    //desc: define o bullet do texto
    //params: (string) tipo do alinhamento
    //return: (obj) Alignment
    static public function type_bullet($type){
        switch ($type){
            case 'TYPE_BULLET':
                return Bullet::TYPE_BULLET;
                break;
            case 'NONE':
            default:
                return Bullet::TYPE_NONE;
        }
    }

    //desc: define a criacao de box como text
    //params: (obj) getActiveSlide, (string) tipo do box
    //return: (obj) createRichTextShape
    static public function type_box($slide, $type){
        switch ($type){
            case 'RICHTEXTSHAPE':
            default:
                return $slide->createRichTextShape();
                break;
        }
    }

    //desc: define o alinhamento do texto
    //params: (string) tipo do alinhamento
    //return: (obj) Alignment
    static public function type_alignment($type){
        switch ($type){
            case 'HORIZONTAL_CENTER':
                return Alignment::HORIZONTAL_CENTER;
                break;
            case 'HORIZONTAL_LEFT':
            default:
                return Alignment::HORIZONTAL_LEFT;
        }
    }

    //desc: cria apresentacao
    //params: nenehum
    //return: (obj) PhpPresentation
    static function new_presentation(){
        return new PhpPresentation();
    }

    //desc: cria slide
    //params: (obj) PhpPresentation, (bool) se primeiro slide
    //return: (obj) getActiveSlide
    static function new_slide($presentation, $first_slide = false){
        //se for o primeiro slide
        if($first_slide) return $presentation->getActiveSlide();
        //demais
        return $presentation->createSlide();
    }

    //desc: cria imagem
    //params: (obj) getActiveSlide, (string) caminho de imagem
    //return: nenhum
    static function set_background_image_in_slide($slide, $file_name){
        $bg_image = new Image();
        $bg_image->setPath(IMAGE_STORAGE."/$file_name");
        $slide->setBackground($bg_image);
    }

    //desc: cria arquivo
    //params: (obj) PhpPresentation, (string) tipo de arquivo, (string) nome do arquivo
    //return
    static function create_pptx_file($presentation, $file_type, $file_name){
        $oWriterPPTX = IOFactory::createWriter($presentation, $file_type); //definindo o tipo de arquivo como PowerPoint2007
        $oWriterPPTX->save(PRESENTATION_STORAGE. "/$file_name");
    }
}