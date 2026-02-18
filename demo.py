from dtool import split_docx, Config, Boundary, Unit, merge_docx

split_docx(Config(
    file_path=r'Moby-Dick_-Or_-The-Whale-by-Herman-Melville.docx',
    output_path='split_output',
    count=90000,
    unit=Unit.WORDS,
    boundary=Boundary.NEAREST_SENTENCE
))

#merge_docx(
#    input_files=r'split_output\output_1771417747984766300',
#    output_path='output_folder',
#    remove_encryption=True
#)