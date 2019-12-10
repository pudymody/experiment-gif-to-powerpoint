use image::gif::Decoder;
use image::AnimationDecoder;
use image::ImageDecoder;
use image::ImageFormat::JPEG;

use std::io::Write;
use std::fs::File;
use std::fs::create_dir;

use std::path::Path;
use walkdir::{WalkDir, DirEntry};
use zip::write::FileOptions;
use std::io::Seek;
use std::io::Read;


fn decode() -> (u64,f64,f64){
    // Decode a gif into frames
    let file_in = File::open("giphy.gif").expect("Could not open file");
    let mut decoder = Decoder::new(file_in).expect("Could not create decoder");

    let dim = decoder.dimensions();
    let frames = decoder.into_frames();
    let frames = frames.collect_frames().expect("error decoding gif");

    let mut i = 0;
    for f in frames {
        println!("{:?}", f.delay().to_integer());
        f.into_buffer().save_with_format(
            format!("dist/media/{}.jpg", i),
            JPEG
        ).expect("Could not save frame");
        i += 1;
    }

    return (i,0.0104166667 * dim.0 as f64,0.0104166667 * dim.1 as f64);
}

fn basicTemplate() {
    create_dir("dist").expect("Could not write template");
    create_dir("dist/media").expect("Could not write template");
    create_dir("dist/META-INF").expect("Could not write template");

    let meta = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
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
    </office:document-meta>";

    let mut file_meta = File::create("dist/meta.xml").expect("Could not write file");
    file_meta.write( meta.as_bytes() ).expect("Could not write meta");

    let mime = "application/vnd.oasis.opendocument.presentation";
    let mut file_mime = File::create("dist/mimetype").expect("Could not write mime");
    file_mime.write( mime.as_bytes() ).expect("Could not write mime");

    let settings = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><office:document-settings xmlns:config='urn:oasis:names:tc:opendocument:xmlns:config:1.0' xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' />";
    let mut file_settings = File::create("dist/settings.xml").expect("Could not write settings");
    file_settings.write( settings.as_bytes() ).expect("Could not write settings");
}

fn processingTemplate( data: (u64, f64, f64) ) {
    let styles = format!("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
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
        <office:automatic-styles>
            <style:page-layout style:name='pageLayout1'>
                <style:page-layout-properties fo:page-width='{}in' fo:page-height='{}in'
                    style:print-orientation='landscape' style:register-truth-ref-style-name='' />
            </style:page-layout>
        </office:automatic-styles>
        <office:master-styles>
            <draw:layer-set>
                <draw:layer draw:name='Master1-bg' draw:protected='true' />
            </draw:layer-set>
            <style:master-page style:name='Master1-Layout1' style:page-layout-name='pageLayout1'
                draw:style-name='a30'>
            </style:master-page>
        </office:master-styles>
    </office:document-styles>", data.1, data.2);
    let mut file_styles = File::create("dist/styles.xml").expect("Could not write file");
    file_styles.write( styles.as_bytes() ).expect("Could not write meta");

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
            <office:automatic-styles>
                <style:style style:family='drawing-page' style:name='a335'>
                    <style:drawing-page-properties draw:fill='solid' draw:fill-color='#ffffff' draw:opacity='100%'
                        presentation:transition-type='automatic' presentation:duration='PT0.400000S' presentation:visibility='visible' draw:background-size='border'
                        presentation:background-objects-visible='true' presentation:background-visible='true'
                        presentation:display-header='false' presentation:display-footer='false'
                        presentation:display-page-number='false' presentation:display-date-time='false' />
                </style:style>
                <style:style style:family='graphic' style:name='a336' style:parent-style-name='Graphics'>
                    <style:graphic-properties draw:fill='none' draw:stroke='none' />
                </style:style>
            </office:automatic-styles>
            <office:body>
                <office:presentation>");

    for i in 0..data.0 {
        let slide = format!("<draw:page draw:name='Slide{0}' draw:style-name='a335'
        draw:master-page-name='Master1-Layout1'
        presentation:presentation-page-layout-name='Master1-PPL1'>
        <draw:frame draw:id='id{0}' draw:style-name='a336' draw:name='{0} Imagen' svg:x='0in'
            svg:y='0in' svg:width='{1}in' svg:height='{2}in' style:rel-width='scale'
            style:rel-height='scale'>
            <draw:image xlink:href='media/{0}.jpg' xlink:type='simple' xlink:show='embed'
                xlink:actuate='onLoad' />
            <svg:desc>trooper3.jpg</svg:desc>
        </draw:frame>
    </draw:page>", i, data.1, data.2);
        content.push_str( slide.as_str() );
    }

    content.push_str("<presentation:settings presentation:endless='true' />
    </office:presentation>
    </office:body>
    </office:document-content>");

    let mut file_content = File::create("dist/content.xml").expect("Could not write file");
    file_content.write( content.as_bytes() ).expect("Could not write meta");

    let mut metainf = String::from("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
    <manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'>
        <manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.presentation' />
        <manifest:file-entry manifest:full-path='META-INF/manifest.xml' manifest:media-type='text/xml' />
        <manifest:file-entry manifest:full-path='content.xml' manifest:media-type='text/xml' />
        <manifest:file-entry manifest:full-path='meta.xml' manifest:media-type='text/xml' />
        <manifest:file-entry manifest:full-path='settings.xml' manifest:media-type='text/xml' />
        <manifest:file-entry manifest:full-path='mimetype' manifest:media-type='text/plain' />
        <manifest:file-entry manifest:full-path='styles.xml' manifest:media-type='text/xml' />");

    for i in 0..data.0 {
        let metafile = format!("<manifest:file-entry manifest:full-path='media/{}.jpg' manifest:media-type='image/jpg' />", i);
        metainf.push_str( metafile.as_str() );
    }

    metainf.push_str("</manifest:manifest>");
    let mut file_manifest = File::create("dist/META-INF/manifest.xml").expect("Could not write file");
    file_manifest.write( metainf.as_bytes() ).expect("Could not write meta");
}

fn zip_dir<T>(it: &mut dyn Iterator<Item=DirEntry>, prefix: &str, writer: T, method: zip::CompressionMethod)
              -> zip::result::ZipResult<()>
    where T: Write+Seek {
    let mut zip = zip::ZipWriter::new(writer);
    let options = FileOptions::default()
        .compression_method(method)
        .unix_permissions(0o755);

    let mut buffer = Vec::new();
    for entry in it {
        let path = entry.path();
        let name = path.strip_prefix(Path::new(prefix)).unwrap();

        // Write file or directory explicitly
        // Some unzip tools unzip files with directory paths correctly, some do not!
        if path.is_file() {
            println!("adding file {:?} as {:?} ...", path, name);
            zip.start_file_from_path(name, options)?;
            let mut f = File::open(path)?;

            f.read_to_end(&mut buffer)?;
            zip.write_all(&*buffer)?;
            buffer.clear();
        } else if name.as_os_str().len() != 0 {
            // Only if not root! Avoids path spec / warning
            // and mapname conversion failed error on unzip
            println!("adding dir {:?} as {:?} ...", path, name);
            zip.add_directory_from_path(name, options)?;
        }
    }
    zip.finish()?;
    Result::Ok(())
}

fn compress(){
    let path = Path::new("presentation.odp");
    let file = File::create(&path).expect("Could not create file");

    let walkdir = WalkDir::new("dist");
    let it = walkdir.into_iter();

    zip_dir(&mut it.filter_map(|e| e.ok()), "dist", file, zip::CompressionMethod::Stored).expect("Could not create file");
}

fn main(){
    basicTemplate();
    let info = decode();
    processingTemplate(info);
    compress();
}