use image::gif::Decoder;
use image::AnimationDecoder;
use image::ImageDecoder;
use image::ImageFormat::JPEG;

use std::io::Write;
use std::io::Seek;
use std::fs::File;
use std::fs::create_dir;
use std::fs::read;
use std::fs::remove_dir_all;

use std::path::Path;
use std::path::PathBuf;
use zip::write::FileOptions;
use zip::ZipWriter;

use std::vec::Vec;
fn decode( src: PathBuf, mut folder: PathBuf) -> Result<(Vec<u16>,f64,f64), String>{
    // Workaround to have some file to work with in the path
    folder.push("0");

    // Decode a gif into frames
    let file_in = File::open( src ).map_err(|_| "Could not open file")?;
    let decoder = Decoder::new(file_in).map_err(|_| "Could not create decoder")?;

    let dim = decoder.dimensions();
    let frames = decoder.into_frames();
    let frames = frames.collect_frames().map_err(|_| "error decoding gif")?;

    let mut delays: Vec<u16> = Vec::new();
    let mut i = 0;
    for f in frames {
        delays.push( f.delay().to_integer() );
        folder.set_file_name( i.to_string() );
        folder.set_extension("jpg");

        f.into_buffer().save_with_format(
            &folder,
            JPEG
        ).map_err(|_| "Could not save frame")?;
        i += 1;
    }

    // const PIXEL_TO_IN:f64 = 0.0104166667;
    return Ok( (delays, dim.0 as f64,dim.1 as f64) );
}

fn basic_template<T>( file: &mut ZipWriter<T> ) -> Result<(), String> where T: Write+Seek {
    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Stored)
        .unix_permissions(0o755);

    file.start_file_from_path( Path::new("meta.xml"), options).map_err(|_| "Could not create presentation")?;
    file.write(b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
    <office:document-meta xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0'
    xmlns:meta='urn:oasis:names:tc:opendocument:xmlns:meta:1.0' xmlns:dc='http://purl.org/dc/elements/1.1/'
    xmlns:xlink='http://www.w3.org/1999/xlink' office:version='1.1'>
    <office:meta>
        <meta:generator>MicrosoftOffice/12.0 MicrosoftPowerPoint</meta:generator>
        <dc:title>Diapositiva 1</dc:title>
        <meta:initial-creator>Creator</meta:initial-creator>
        <dc:creator>Creator</dc:creator>
        <meta:creation-date>2019-12-09T02:44:17Z</meta:creation-date>
        <dc:date>2019-12-09T02:58:24Z</dc:date>
        <meta:editing-cycles>1</meta:editing-cycles>
        <meta:editing-duration>PT0S</meta:editing-duration>
        <meta:document-statistic meta:paragraph-count='0' meta:word-count='0' />
    </office:meta>
    </office:document-meta>").map_err(|_| "Could not create presentation")?;

    file.start_file_from_path( Path::new("mimetype"), options).map_err(|_| "Could not create presentation")?;
    file.write(b"application/vnd.oasis.opendocument.presentation").map_err(|_| "Could not create presentation")?;

    file.start_file_from_path( Path::new("settings.xml"), options).map_err(|_| "Could not create presentation")?;
    file.write(b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?><office:document-settings xmlns:config='urn:oasis:names:tc:opendocument:xmlns:config:1.0' xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' />")
        .map_err(|_| "Could not create presentation")?;

    return Ok(());
}

fn processing_template<T>( file: &mut ZipWriter<T>, data: (Vec<u16>, f64, f64) ) -> Result<(), String> where T: Write+Seek {
    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Stored)
        .unix_permissions(0o755);

    file.start_file_from_path( Path::new("styles.xml"), options).map_err(|_| "Could not create presentation")?;
    file.write(
        format!("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
            <office:document-styles xmlns:dom='http://www.w3.org/2001/xml-events'
            xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'
            xmlns:fo='urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
            xmlns:presentation='urn:oasis:names:tc:opendocument:xmlns:presentation:1.0'
            xmlns:smil='urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0'
            xmlns:style='urn:oasis:names:tc:opendocument:xmlns:style:1.0'
            xmlns:svg='urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0'
            xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0'
            xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:xlink='http://www.w3.org/1999/xlink'
            xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0'>
            <office:styles/>
            <office:automatic-styles>
                <style:page-layout style:name='pageLayout1'>
                    <style:page-layout-properties fo:page-width='{}px' fo:page-height='{}px'
                        style:print-orientation='landscape' />
                </style:page-layout>
            </office:automatic-styles>
            <office:master-styles>
                <draw:layer-set>
                    <draw:layer draw:name='Master1-bg' draw:protected='true' />
                </draw:layer-set>
                <style:master-page style:name='Master1' style:page-layout-name='pageLayout1'>
                </style:master-page>
            </office:master-styles>
        </office:document-styles>", data.1, data.2).as_bytes()
    ).map_err(|_| "Could not create presentation")?;

    let mut content = String::from("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
        <office:document-content xmlns:dom='http://www.w3.org/2001/xml-events'
            xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0'
            xmlns:fo='urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
            xmlns:presentation='urn:oasis:names:tc:opendocument:xmlns:presentation:1.0'
            xmlns:script='urn:oasis:names:tc:opendocument:xmlns:script:1.0'
            xmlns:smil='urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0'
            xmlns:style='urn:oasis:names:tc:opendocument:xmlns:style:1.0'
            xmlns:svg='urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0'
            xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0'
            xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:xlink='http://www.w3.org/1999/xlink'
            xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0'>
            <office:automatic-styles>");

    let mut slide_style: String = String::new();
    let mut slide_content: String = String::new();
    let mut metainf = String::from("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
        <manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'>
        <manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.presentation' />
        <manifest:file-entry manifest:full-path='META-INF/manifest.xml' manifest:media-type='text/xml' />
        <manifest:file-entry manifest:full-path='content.xml' manifest:media-type='text/xml' />
        <manifest:file-entry manifest:full-path='meta.xml' manifest:media-type='text/xml' />
        <manifest:file-entry manifest:full-path='settings.xml' manifest:media-type='text/xml' />
        <manifest:file-entry manifest:full-path='mimetype' manifest:media-type='text/plain' />
        <manifest:file-entry manifest:full-path='styles.xml' manifest:media-type='text/xml' />");

    for (i,f) in data.0.iter().enumerate() {
        let slide_style_item = format!("
            <style:style style:family='drawing-page' style:name='a{0}'>
                <style:drawing-page-properties draw:fill='solid' draw:fill-color='#ffffff' draw:opacity='100%'
                    presentation:transition-type='automatic' presentation:duration='PT{1:.3}S' presentation:visibility='visible' draw:background-size='border'
                    presentation:background-objects-visible='true' presentation:background-visible='true'
                    presentation:display-header='false' presentation:display-footer='false'
                    presentation:display-page-number='false' presentation:display-date-time='false' />
            </style:style>",
            i, (*f as f64) / 100.0);

        let slide = format!("<draw:page draw:name='Slide{0}' draw:style-name='a{0}'
            draw:master-page-name='Master1'>
                <draw:frame draw:id='id{0}' draw:style-name='a336' draw:name='{0} Imagen' svg:x='0px'
                    svg:y='0px' svg:width='{1}px' svg:height='{2}px' style:rel-width='scale'
                    style:rel-height='scale'>
                    <draw:image xlink:href='media/{0}.jpg' xlink:type='simple' xlink:show='embed'
                        xlink:actuate='onLoad' />
                    <svg:desc>trooper3.jpg</svg:desc>
                </draw:frame>
            </draw:page>",
            i, data.1, data.2);

        let metafile = format!("<manifest:file-entry manifest:full-path='media/{}.jpg' manifest:media-type='image/jpg' />", i);

        slide_content.push_str( slide.as_str() );
        slide_style.push_str( slide_style_item.as_str() );
        metainf.push_str( metafile.as_str() );
    }

    content.push_str( &slide_style );
    content.push_str("
        <style:style style:family='graphic' style:name='a336' style:parent-style-name='Graphics'>
            <style:graphic-properties draw:fill='none' draw:stroke='none' />
        </style:style>
    </office:automatic-styles>
    <office:body>
        <office:presentation>");

    content.push_str( &slide_content );
    content.push_str("<presentation:settings presentation:endless='true' />
    </office:presentation>
    </office:body>
    </office:document-content>");

    file.start_file_from_path( Path::new("content.xml"), options).map_err(|_| "Could not create presentation")?;
    file.write( content.as_bytes() ).map_err(|_| "Could not create presentation")?;

    metainf.push_str("</manifest:manifest>");
    file.start_file_from_path( Path::new("META-INF/manifest.xml"), options).map_err(|_| "Could not create presentation")?;
    file.write( metainf.as_bytes() ).map_err(|_| "Could not create presentation")?;

    return Ok(());
}

fn main() -> Result<(), String>{
    let presentation_file = File::create("presentation.odp").map_err(|_| "Could not create presentation file")?;
    let mut zip = ZipWriter::new(presentation_file);

    basic_template(&mut zip)?;

    create_dir("dist").map_err(|_| "Could not write template")?;
    create_dir("dist/media").map_err(|_| "Could not write template")?;
    let info = decode( PathBuf::from("image.gif"), PathBuf::from("dist/media/") )?;

    let options = FileOptions::default()
        .compression_method(zip::CompressionMethod::Stored)
        .unix_permissions(0o755);

    let mut frame_path_in_zip = PathBuf::from("media/0.jpg");
    let mut frame_path = PathBuf::from("dist/media/0.jpg");
    for i in 0..info.0.len() {
        frame_path_in_zip.set_file_name( i.to_string() );
        frame_path_in_zip.set_extension("jpg");
        zip.start_file_from_path( &frame_path_in_zip, options).map_err(|_| "Could not save frame")?;

        frame_path.set_file_name( i.to_string() );
        frame_path.set_extension("jpg");
        let frame_content = read( &frame_path ).map_err(|_| "Could not save frame")?;
        zip.write( &frame_content ).map_err(|_| "Could not save frame")?;
    }

    remove_dir_all( PathBuf::from("dist/") ).map_err(|_| "Could not clean frames")?;

    processing_template(&mut zip, info)?;

    return Ok(());
}