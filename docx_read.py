import docx
import argparse
import pandas as pd

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
                    description = 'Takes a docx with an exam in format 1.-Question\na.-answer\nb.-answer and returns a csv file with the format Question;answer1;answer2; ...;answerN to be the input for a main_docx.py',
                    )
    parser.add_argument('-i', '--input', required=True, help="The file with the exam in dx")
    parser.add_argument('-o', '--output', required=True, help="The csv file to be generated")
    args = vars(parser.parse_args())

    # open connection to Word Document
    doc = docx.Document(args['input'])
    if doc is None:
        print(f'Error opening file {args["input"]}')
        exit(1)

    # read in each paragraph in file
    result = [p.text for p in doc.paragraphs]
    i=0
    csv_content=[]
    while i < len(result):
        if isinstance(result[i], str) and len(result[i].strip())>0:
            j = 1
            question = {'question': result[i].strip()}
            while isinstance(result[i+j], str) and len(result[i+j].strip())>0:
                question[f'answer_{j}'] = result[i+j].strip()
                j+=1
            i = i+j
            csv_content.append(question)
        else:
            i+=1
    df = pd.DataFrame(csv_content)
    df.to_csv(args['output'], sep=';', index=False)
    print(f'Generated file {args["output"]}')