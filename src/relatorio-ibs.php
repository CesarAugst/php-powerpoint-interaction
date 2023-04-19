<?php
/*Bibliotecas */
//utiliza bibliotecas do composer
require_once 'vendor/autoload.php';
//bibliotecas envolvidas
use PhpOffice\PhpPresentation\PhpPresentation; //classe do PhpPresentation
use PhpOffice\PhpPresentation\Slide\Background\Image; //utilizacao de imagens
use PhpOffice\PhpPresentation\IOFactory; //classe para manipular os arquivos
use PhpOffice\PhpPresentation\Style\Alignment; //classe de estilo de alinhamentos
use PhpOffice\PhpPresentation\Style\Color; //classe de estilo de cores

/*Constantes de apoio*/
//files
const IMAGE_STORAGE = __DIR__ . './assets/images';//diretorio das imagens
const PRESENTATION_STORAGE = __DIR__ . './assets/presentation_files'; //diretorio dos arquivos de apresentacao
//style
const TITLE_PRIMARY_COLOR = "FFFFFF";


/*Criacao */
//inicia nova apresentacao
$presentation = new PhpPresentation();

//cria slide
$slide_1 = $presentation->getActiveSlide();

// Slide > Background > Image
$bg_image = new Image();
$bg_image->setPath(IMAGE_STORAGE."/ibs-bg.png");
$slide_1->setBackground($bg_image);

//titulo_capa
create_text(
    $slide_1, //slide
    300, //altura
    600, //largura
    20, //posicao no eixo X
    400, //posicao no eixo Y
    "Relatório de Imprensa", //texto
    true, //ativo bold
    40, //font size
    TITLE_PRIMARY_COLOR //cor da fonte
);
//subtiulo

//salva arquivo
$oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007'); //definindo o tipo de arquivo como PowerPoint2007
$oWriterPPTX->save(PRESENTATION_STORAGE. "/teste-ibs.pptx");

//desc: criacao de texto
//params:
//return: (obj) RichTextShape
function create_text($slide, $height, $width, $offsetX, $offsetY, $text, $isBold, $fontSize, $color){
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