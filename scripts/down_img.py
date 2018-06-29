import math
import os
import requests
import threading

import xlrd
import xlwt

PROJ_ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE_DIR = os.path.join(PROJ_ROOT_DIR, 'files')
IMG_DIR = os.path.join(FILE_DIR, 'images')
TOUTIAO_ANNO_FILE_PATH = os.path.join(FILE_DIR, 'toutiao_content_image.xls')
OUTPUT_XLS_FILE_PATH = os.path.join(FILE_DIR, 'article_attr.xls')
TOUTIAO_ANNO_TITLES = ('article_title', 'article_class', 'tags', 'article_url',
                       'img_urls', 'article_content')
OUTPUT_XLS_TITLES = ('article_id', 'img_urls', 'img_paths', 'has_content',
                     'contain_car', 'contain_human', 'human_tag',
                     'object tags')


def down_img(url, save_path):
    try:
        rsp = requests.get(url)
        if rsp.status_code == 200:
            with open(save_path, 'wb') as f:
                f.write(rsp.content)
        return True

    except Exception as e:
        print(e)
        return False


def process_article(output_xls_sheet, article_id, row_value):
    """Performs image downloading and data processing for an article.

    :param output_xls_sheet: The xls sheet to write (xlwt.Workbook.Sheet).
    :param article_id: Index for the article in the anno file (int).
    :param row_value: The data for that article in the anno file (list).
    :returns: None
    :rtype: None

    """
    article_class = row_value[1].strip()
    img_urls = row_value[4].strip()
    article_content = row_value[-1].strip()

    img_url_list = img_urls.split(',') if img_urls else []
    img_path_list = []
    has_content = (article_content != '')
    contain_car = '汽车' in article_class
    article_dir = os.path.join(IMG_DIR, 'article_' + str(article_id))
    non_prefix_article_dir = os.path.join(
        os.path.basename(IMG_DIR), 'article_' + str(article_id))

    if not os.path.exists(article_dir):
        os.makedirs(article_dir)

    for i, img_url in enumerate(img_url_list):
        img_save_path = os.path.join(article_dir, 'img_' + str(i) + '.jpg')
        non_prefix_save_path = os.path.join(non_prefix_article_dir,
                                            'img_' + str(i) + '.jpg')

        if os.path.exists(img_save_path):
            existp = True
        else:
            existp = down_img(img_url, img_save_path)

        img_path_list.append(non_prefix_save_path if existp else '')

    img_paths = ','.join(img_path_list)

    # article_id + 1 because the first row of the sheet is the tile row
    output_xls_sheet.write(article_id + 1, 0, article_id)
    output_xls_sheet.write(article_id + 1, 1, img_urls)
    output_xls_sheet.write(article_id + 1, 2, img_paths)
    output_xls_sheet.write(article_id + 1, 3, int(has_content))
    output_xls_sheet.write(article_id + 1, 4, int(contain_car))
    output_xls_sheet.write(article_id + 1, 5, '')
    output_xls_sheet.write(article_id + 1, 6, '')
    output_xls_sheet.write(article_id + 1, 7, '')


def process_split(split, input_xls_sheet, output_xls_sheet):
    """Processes a list of articles in the split.

    :param split: The indices of the articles (list).
    :param input_xls_sheet: The xls sheet to read from.
    :param output_xls_sheet: The xls sheet to write.
    :returns: None.
    :rtype: None.

    """
    for article_id in split:
        row_value = input_xls_sheet.row_values(article_id)
        process_article(output_xls_sheet, article_id, row_value)


def main():
    # prepare the input and output excel files
    toutiao_anno_file = xlrd.open_workbook(TOUTIAO_ANNO_FILE_PATH)
    toutiao_anno_sheet = toutiao_anno_file.sheet_by_index(0)
    output_xls_file = xlwt.Workbook()
    output_xls_sheet = output_xls_file.add_sheet(
        'toutiao', cell_overwrite_ok=True)

    for i, title in enumerate(OUTPUT_XLS_TITLES):
        output_xls_sheet.write(0, i, title)

    # TODO: remove '[:100]' to access the full list
    article_ids = list(range(0, toutiao_anno_sheet.nrows))[:100]
    n_splits = 8
    split_size = math.ceil(len(article_ids) / n_splits)
    splits = []

    for i_split in range(n_splits):
        start = i_split * split_size
        end = (i_split + 1) * split_size
        if end > len(article_ids):
            end = len(article_ids)
        split = article_ids[start:end]
        splits.append(split)

    threads = []
    for split in splits:
        thread = threading.Thread(
            target=process_split,
            args=(split, toutiao_anno_sheet, output_xls_sheet))
        thread.start()
        threads.append(thread)

    for thread in threads:
        thread.join()

    output_xls_file.save(OUTPUT_XLS_FILE_PATH)


if __name__ == '__main__':
    main()
