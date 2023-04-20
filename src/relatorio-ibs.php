<?php
/*Bibliotecas */
//utiliza bibliotecas do composer
require_once 'vendor/autoload.php';
require_once 'assets/class/LibPhpPresentationManipulation.php';

/*Constantes de apoio*/
//files
const IMAGE_STORAGE = __DIR__ . './assets/images';//diretorio das imagens
const PRESENTATION_STORAGE = __DIR__ . './assets/presentation_files'; //diretorio dos arquivos de apresentacao
//style
const TITLE_PRIMARY_COLOR = "FFFFFF";
const TITLE_SECONDARY_COLOR = "FF000000";
//alignment
const HORIZONTAL_CENTER = "HORIZONTAL_CENTER";
const HORIZONTAL_LEFT = "HORIZONTAL_LEFT";
//bullet
const TYPE_BULLET = "TYPE_BULLET";
const TYPE_NONE = "TYPE_NONE";
//type box
const RICHTEXTSHAPE = "RICHTEXTSHAPE";

/*Instancia da manipulacao*/
$lib_pptx = new LibPhpPresentationManipulation();

/*Criacao */
//inicia nova apresentacao
$presentation = $lib_pptx::new_presentation();

//cria slide
$slide_1 = $lib_pptx::new_slide($presentation, true);
// Slide > Background > Image
$lib_pptx::set_background_image_in_slide(
    $slide_1, //slide
    "ibs-bg-primary.png" //nome do arquivo
);
//cria box para o titulo
$created_box = $lib_pptx::create_box(
    $lib_pptx::type_box($slide_1, RICHTEXTSHAPE), //tipo de box
    300, //altura
    600, //largura
    60, //posicao no eixo X
    350 //posicao no eixo Y
);
//titulo_capa
$lib_pptx::create_text(
    $created_box, //box do texto
    $lib_pptx::type_alignment(HORIZONTAL_LEFT), //alinhamento texto
    "Relatório de Imprensa", //texto
    true, //ativo bold
    45, //font size
    TITLE_PRIMARY_COLOR //cor da fonte
);
//cria box para o subtitulo
$created_box = $lib_pptx::create_box(
    $lib_pptx::type_box($slide_1, RICHTEXTSHAPE), //tipo de box
    300, //altura
    600, //largura
    60, //posicao no eixo X
    450 //posicao no eixo Y
);
//subtitulo_capa
$lib_pptx::create_text(
    $created_box, //tipo de box
    $lib_pptx::type_alignment(HORIZONTAL_LEFT), //alinhamento do texto
    "Março 2023", //texto
    false, //ativo bold
    30, //font size
    TITLE_PRIMARY_COLOR //cor da fonte
);

//cria slide
$slide_2 = $lib_pptx::new_slide($presentation);
// Slide > Background > Image
$lib_pptx::set_background_image_in_slide(
    $slide_2, //slide
    "ibs-bg-secondary.png" //nome do arquivo
);
//cria box para o titulo
$created_box = $lib_pptx::create_box(
    $lib_pptx::type_box($slide_2, RICHTEXTSHAPE), //tipo de box
    300, //altura
    600, //largura
    50, //posicao no eixo X
    40 //posicao no eixo Y
);
//titulo
$lib_pptx::create_text(
    $created_box, //box do texto
    $lib_pptx::type_alignment(HORIZONTAL_LEFT), //alinhamento do texto
    "Atividades Desenvolvidas", //texto
    true, //ativo bold
    30, //font size
    TITLE_SECONDARY_COLOR //cor da fonte
);
//cria box para o paragrafo
$created_box = $lib_pptx::create_box(
    $lib_pptx::type_box($slide_2, RICHTEXTSHAPE), //tipo de box
    300, //altura
    600, //largura
    50, //posicao no eixo X
    150 //posicao no eixo Y
);
//paragrafo
$lib_pptx::create_paragraph_text(
    $created_box, //box
    $lib_pptx::type_alignment(HORIZONTAL_LEFT), //alinhamento do texto
    ["Assessoria de Imprensa;", "Monitoramento de mercado;", "Cobertura de eventos;", "Ações de Relacionamento com a mídia."], //texto
    false, //ativo bold
    20, //font size
    TITLE_SECONDARY_COLOR, //cor da fonte
    $lib_pptx::type_bullet(TYPE_BULLET) //tipo de bullet
);
//cria box para o numero
$created_box = $lib_pptx::create_box(
    $lib_pptx::type_box($slide_2, RICHTEXTSHAPE), //tipo de box
    10, //altura
    10, //largura
    920, //posicao no eixo X
    680 //posicao no eixo Y
);
//numero da pagina
$lib_pptx::create_text(
    $created_box, //box do texto
    $lib_pptx::type_alignment(HORIZONTAL_CENTER), //alinhamento do texto
    "2", //texto
    false, //ativo bold
    20, //font size
    TITLE_PRIMARY_COLOR //cor da fonte
);


//salva arquivo
$lib_pptx::create_pptx_file(
    $presentation, //apresentacao
    'PowerPoint2007', //tipo de arquivo
    "teste-ibs.pptx"
);

