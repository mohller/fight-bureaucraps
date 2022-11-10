from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject, TextStringObject, NumberObject, encode_pdfdocencoding
from PyPDF2.constants import FieldFlag, PageAttributes, AnnotationDictionaryAttributes, InteractiveFormDictEntries, StreamAttributes, FilterTypes, FieldDictionaryAttributes
from PyPDF2.filters import FlateDecode


"""This code was sketched from the answers to the question 
https://stackoverflow.com/questions/47288578/pdf-form-filled-with-pypdf2-does-not-show-in-print

   The code for ticking the checkboxes was inspired in the answer to 
https://stackoverflow.com/questions/35538851/how-to-check-uncheck-checkboxes-in-a-pdf-with-python-preferably-pypdf2
"""

user_settings = {
    'Name Vorname' : 'Paco Manolo',
    'Amtsbezeichnung' : 'Doktorand',
    'TelNr' : '-3736',
    'EmailAdresse' : 'paco.manolo@uni-wuppertal.de',
    'PLZ  Wohnort': '42561',
    'FakultätZEDezernatsonst': 'Fk4',
    'Dienstort' : 'Wuppertal',
    ' Klasse' : '1st'
}

unknown_fields = [
    '0_4',
    'Text2',
    'Text3',
    'Text4',
    'Text5',
    'Text6',
    'Text7',
    'Text8',
    'Text9',
    'Text10',
    'Text11'
]

ForeignDestination = True
kostenstelle = 'C0207101A'
input_data = dict(((fn, va) for va, fn in zip(kostenstelle, unknown_fields[1:-1])))
input_data[unknown_fields[-1]] = '0'
input_data[unknown_fields[0]] = str(int(ForeignDestination))
input_data.update(user_settings)

travel_settings = {
    'Ziele der Reise OrtLand' : 'Barcelona, Spain',
    'Zweck der Reise' : 'International Conference of Researchers Named Manolo',
    'Beginn der Reise DatumUhrzeit' : '01.04.2025 08:00',
    'Beginn des Dienstgeschäftes DatumUhrzeit':'20.04.2025 16:20',
    'Ende der Reise DatumUhrzeit' : '20.04.2025 16:24',
    'Begründung Dienstwagen' : 'is cheaper',
    'Begründung Mietwagen' : 'is cheaper',
    'Begründung Flugzeug' : 'is cheaper',
    'Fahr-und Flugkosten vorauss' : '10eur',
    'ÜK vorauss' : '10eur',
    'TNG vorauss' : '10eur',
    'sonstiges vorauss' : '10eur',
    ' Reisebeihilfe' : 'd',
    'Einschränkung Genehmigung' : 'd'
}

other_settings = {
    'Ich besitze KEINE BahnCard' : 'On'
}

input_data.update(travel_settings)
input_data.update(other_settings)

tick_fields = [
    'Ich besitze eine BahnCard 100',
    'Ich besitze eine BahnCard 50',
    'Ich besitze eine BahnCard Business 25',
    'Ich besitze eine BahnCard Business 50',
    'Öffentliche Verkehrsmittel 2Klasse unter Nutzung des Firmenkundenrabatts Nr 5000317',
    'Öffentliche Verkehrsmittel 1 Klasse unter Nutzung des Firmenkundenrabatts Nr 5000317',
    'Dienstwagen  Begründung'
]

# --------------------------------------
# Filling data.
data = input_data

def load_template_data(source):
    """Read data from a form previously filled and returns a
    dictionary with the values contained and the keys are the
    ids of the inputs of the html.
    """
    pass


def create_travel_form_pdf(template_filename):
    """Takes data from html and writes it into a new file based
    on the template.
    """
    
    # Get template
    template = PdfReader(template_filename, strict=False)

    if "/AcroForm" in template.trailer["/Root"]:
        template.trailer["/Root"]["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)}
        )

    # Initialize writer.
    writer = PdfWriter()
    # set_need_appearances_writer(writer)
    if "/AcroForm" in writer._root_object:
        writer._root_object["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)}
        )

    # Add the template page.
    writer.add_page(template.pages[0])

    # Get page annotations.
    page_annotations = writer.pages[0][PageAttributes.ANNOTS]

    # Loop through page annotations (fields).
    for page_annotation in page_annotations:  # type: ignore
        # Get annotation object.
        annotation = page_annotation.get_object()  # type: ignore

        # Get existing values needed to create the new stream and update the field.
        field = annotation.get(NameObject("/T"))
        if field in tick_fields:
            annotation.update({
                NameObject("/V") : NameObject("/On"),
                NameObject("/AS"): NameObject("/On")
            })
        new_value = data.get(field, 'True')
        
        ap = annotation.get(AnnotationDictionaryAttributes.AP)
        try:
            x_object = ap.get(NameObject("/N")).get_object()
        except:
            print('xobject creation did not work with Fieldname:', field)
            continue

        font = annotation.get(InteractiveFormDictEntries.DA)
        
        rect = annotation.get(AnnotationDictionaryAttributes.Rect)
        
        # Calculate the text position.
        try:
            posTf = font.split(" ").index("Tf")
        except:
            print('font splitting did not work with Fieldname:', field)
            continue

        font_size = float(font.split(" ")[posTf - 1])
        w = round(float(rect[2] - rect[0] - 2), 2)
        h = round(float(rect[3] - rect[1] - 2), 2)
        text_position_h = h / 2 - font_size / 3  # approximation

        # Create a new XObject stream.
        new_stream = f'''
            /Tx BMC 
            q
            1 1 {w} {h} re W n
            BT
            {font}
            2 {text_position_h} Td
            ({new_value}) Tj
            ET
            Q
            EMC
        '''

        # Add Filter type to XObject.
        x_object.update(
            {
                NameObject(StreamAttributes.FILTER): NameObject(FilterTypes.FLATE_DECODE)
            }
        )

        # Update and encode XObject stream.
        x_object._data = FlateDecode.encode(encode_pdfdocencoding(new_stream))
        if field in ['Ich besitze KEINE BahnCard', 'Ich besitze eine BahnCard 25', 'TelNr']:
            print(x_object)

        # Update annotation dictionary.
        annotation.update(
            {
                # Update Value.
                NameObject(FieldDictionaryAttributes.V): TextStringObject(
                    new_value   
                ),
                # Update Default Value.
                NameObject(FieldDictionaryAttributes.DV): TextStringObject(
                    new_value
                ),
                # Set Read Only flag.
                # NameObject(FieldDictionaryAttributes.Ff): NumberObject(
                #     FieldFlag(3)
                # )
            }, 
        )


    # Clone document root & metadata from template.
    # This is required so that the document doesn't try to save before closing.
    writer.clone_reader_document_root(template)

    # write "output".
    with open(f"/home/leonel/Documents/TravelForms/aaa_filled-out.pdf", "wb") as output_stream:
        writer.write(output_stream)  # type: ignore