from dtool import split_docx, Config, Boundary, Unit, merge_docx


def split():
    split_docx(
        Config(
            file_path=r"Moby-Dick_-Or_-The-Whale-by-Herman-Melville.docx",
            output_path="split_output",
            count=10000,
            unit=Unit.CHARS,
            boundary=Boundary.STRICT,
        )
    )


def merge():
    merge_docx(
        input_files=r"split_output", output_path="output_folder", remove_encryption=True
    )


if __name__ == "__main__":
    split()
    # merge()
