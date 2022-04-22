from os import environ
from tempfile import NamedTemporaryFile

from telegram.ext import (
    Updater,
    CommandHandler,
    MessageHandler,
    Filters,
    ConversationHandler,
    DictPersistence,
)

from excel import open_excel_file, find_diff

WAITING_FIRST_FILE, WAITING_SECOND_FILE = range(2)

ASK_FIRST_FILE_MESSAGE = "Надішліть перший файл для порівняння формату xlsx/xls"
ASK_SECOND_FILE_MESSAGE = "Надішліть перший файл для порівняння формату xlsx/xls"
INVALID_MESSAGE_WAITING_FOR_FILE = "Надішліть файл формату xlsx/xls або скасуйте порівняння командою /cancel"
INVALID_MESSAGE_WAITING_START = "Для початку порівняння використайте команду /compare"
CANCEL_MESSAGE = "Порівняння скасовано"


def start(update, context):
    update.message.reply_text(ASK_FIRST_FILE_MESSAGE)
    return WAITING_FIRST_FILE


def invalid_first_file(update, context):
    update.message.reply_text(INVALID_MESSAGE_WAITING_FOR_FILE)
    return WAITING_FIRST_FILE


def process_first_file(update, context):
    file = context.bot.get_file(update.message.document)
    context.user_data['first_file'] = file
    update.message.reply_text(ASK_SECOND_FILE_MESSAGE)
    return WAITING_SECOND_FILE


def invalid_second_file(update, context):
    update.message.reply_text(INVALID_MESSAGE_WAITING_FOR_FILE)
    return WAITING_SECOND_FILE


def compare(update, context):
    chat_id = update.message.chat_id
    first = context.user_data['first_file']
    second = context.bot.get_file(update.message.document)

    first_file_name = first.file_path.split('/')[-1]
    second_file_name = second.file_path.split('/')[-1]

    with NamedTemporaryFile(prefix='compared_', suffix=f'_{first_file_name}') as first_tmp, \
            NamedTemporaryFile(suffix=second_file_name) as second_tmp:
        first.download(out=first_tmp)
        second.download(out=second_tmp)
        first_excel_file = open_excel_file(first_tmp.name)
        second_excel_file = open_excel_file(second_tmp.name)
        highlights = find_diff(first_excel_file, second_excel_file)
        first_excel_file.copy_with_highlights(first_tmp.name, highlights)
        first_tmp.seek(0)
        context.bot.send_document(chat_id, first_tmp)
    update.message.reply_text(
        "\n".join(
            ['Кількість клітинок, що відрізняються:'] +
            [f"{sheet}: {len(cells)}"
             for sheet, cells in highlights.items()]
        )
    )
    return ConversationHandler.END


def cancel(update, context):
    context.user_data.pop('first_file', None)
    update.message.reply_text(CANCEL_MESSAGE)
    return ConversationHandler.END


def waiting_start(update, context):
    update.message.reply_text(INVALID_MESSAGE_WAITING_START)


persistence = DictPersistence()


def main():
    token = environ.get('TELEGRAM_BOT_TOKEN')
    if not token:
        exit(1)
    updater = Updater(token, persistence=persistence)

    dispatcher = updater.dispatcher

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler('compare', start)],
        states={
            WAITING_FIRST_FILE: [
                MessageHandler(
                    Filters.document.file_extension('xls') | Filters.document.file_extension('xlsx'), process_first_file
                ),
                MessageHandler(Filters.all, invalid_first_file),
            ],
            WAITING_SECOND_FILE: [
                MessageHandler(
                    Filters.document.file_extension('xls') | Filters.document.file_extension('xlsx'), compare
                ),
                MessageHandler(Filters.all, invalid_second_file),
            ],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
        name='Compare Conversation',
        persistent=True,
    )

    dispatcher.add_handler(conv_handler)
    dispatcher.add_handler(MessageHandler(
        Filters.all, lambda u, _: u.message.reply_text(INVALID_MESSAGE_WAITING_START),
    ))

    updater.start_polling()

    updater.idle()


if __name__ == '__main__':
    main()
