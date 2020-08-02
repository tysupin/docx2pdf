from win32com import client
import pythoncom
from pathlib import Path
from tqdm.auto import tqdm


def resolve_paths(input_path, output_path):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve() if output_path else None
    output = {}
    if input_path.is_dir():
        output["batch"] = True
        output["input"] = str(input_path)
        if output_path:
            assert output_path.is_dir()
        else:
            output_path = str(input_path)
        output["output"] = output_path
    else:
        output["batch"] = False
        assert str(input_path).endswith(".doc")
        output["input"] = str(input_path)
        if output_path and output_path.is_dir():
            output_path = str(output_path / (str(input_path.stem) + ".pdf"))
        elif output_path:
            assert str(output_path).endswith(".pdf")
        else:
            output_path = str(input_path.parent / (str(input_path.stem) + ".pdf"))
        output["output"] = output_path
    return output

"""
    : word file to pdf
    :param input_path word file name or dir
    :param output_path The name of the converted pdf file or dir
"""

def convert(input_path, output_path=None, keep_active=False):

    word = client.Dispatch("Word.Application")
    wdFormatPDF = 17

    paths = resolve_paths(input_path, output_path)

    if paths["batch"]:
        for doc_filepath in tqdm(sorted(Path(paths["input"]).glob("*.doc"))):
            print("\n"+str(doc_filepath))
            pdf_filepath = Path(paths["output"]) / (str(doc_filepath.stem) + ".pdf")
            doc = word.Documents.Open(str(doc_filepath))
            doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
            doc.Close()
    else:
        pbar = tqdm(total=1)
        doc_filepath = Path(paths["input"]).resolve()
        pdf_filepath = Path(paths["output"]).resolve()
        doc = word.Documents.Open(str(doc_filepath))
        doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
        doc.Close()
        pbar.update(1)

    if not keep_active:
        word.Quit()



#convert("D:\\")
#convert("D:\\test.doc", "D:\\test.pdf")