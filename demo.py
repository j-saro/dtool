from dtool import split_docx, Config, Boundary, Unit, merge_docx


def split():
    split_docx(
        Config(
            file_path=r"Moby-Dick_-Or_-The-Whale-by-Herman-Melville.docx",
            output_path="split_output",
            count=90000,
            unit=Unit.CHARS,
            boundary=Boundary.NEAREST_SENTENCE,
        )
    )


def merge_folder():
    merge_docx(
        input_files=r"split_output", output_path="merge_output", remove_encryption=True
    )

def merge_files():
    files = [
            r"split_output\0002_Moby-Dick_-Or_-The-Whale-by-Herman-Melville.docx",
            r"split_output\0001_Moby-Dick_-Or_-The-Whale-by-Herman-Melville.docx"
        ]
    merge_docx(
        input_files=files, output_path="merge_output", remove_encryption=True
    )

if __name__ == "__main__":
    split()
    merge_folder()
    merge_files()
