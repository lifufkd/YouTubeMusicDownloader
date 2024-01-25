import openpyxl
import pandas as pd
import youtube_dl


def _get_link_if_exists(cell) -> str | None:
    try:
        return cell.hyperlink.target
    except AttributeError:
        return None


def Download(data: list):
    print(data)
    for desc, video_url in data:
        try:
            video_info = youtube_dl.YoutubeDL().extract_info(url = video_url,download=False)
            filename = f"Lost Wave/{desc}.mp3"
            options={
                'format':'bestaudio/best',
                'keepvideo':False,
                'outtmpl':filename,
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                }],
            }
            with youtube_dl.YoutubeDL(options) as ydl:
                ydl.download([video_info['webpage_url']])
        except:
            pass


def extract_hyperlinks_from_xlsx(
    file_name: str, sheet_name: str, columns_to_parse: list[str],  row_header: int = 1
) -> pd.DataFrame:
    temp_data = []
    df = pd.read_excel(file_name, sheet_name)
    ws = openpyxl.load_workbook(file_name)[sheet_name]
    row_offset = row_header + 1
    column_index = list(df.columns).index(columns_to_parse[0]) + 1
    df[columns_to_parse[0]] = [_get_link_if_exists(ws.cell(row=row_offset + i, column=column_index)) for i in range(len(df[columns_to_parse[0]]))]
    for i, g in zip(df[columns_to_parse[0]], df[columns_to_parse[1]]):
        if i is not None:
            if '/' in g:
                x = g.replace('/', 'or')
            else:
                x = g
            temp_data.append([x, i])
    Download(temp_data)


print(extract_hyperlinks_from_xlsx('1.xlsx', 'Mysterious Songs', ['Link', "Placeholder Title"]))