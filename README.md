# pdf-to-word-backend

Конвертирует PDF в .docx через Claude API (vision-разбор каждой страницы: текст, заголовки, таблицы, с пометкой неуверенных фрагментов).

## Как это работает

1. `POST /convert` принимает PDF (multipart, поле `file`) и опционально `pages` — список номеров страниц через запятую (`"1,3,5"`), если нужны не все.
2. Каждая страница рендерится в JPEG через `pdftoppm` (poppler-utils).
3. Каждая картинка отправляется в Claude (`messages.create` с vision) с промптом, который просит вернуть JSON-структуру блоков (заголовки/абзацы/таблицы), помечая неуверенные фрагменты `{{...}}`.
4. Результат собирается в .docx (`docx` npm-пакет) — неуверенные фрагменты подсвечиваются цветом прямо в документе.
5. Сервер отдаёт готовый .docx файл в ответе.

## Локальный запуск

```bash
cp .env.example .env
# впиши свой ANTHROPIC_API_KEY в .env

npm install
npm start
```

Сервер поднимется на `http://localhost:3000`. Проверка: `GET /health` → `{"ok":true}`.

**Системная зависимость:** нужен `poppler-utils` (даёт бинарник `pdftoppm`).
- Ubuntu/Debian: `sudo apt install poppler-utils`
- macOS: `brew install poppler`

## Пример запроса

```bash
curl -X POST http://localhost:3000/convert \
  -F "file=@/path/to/document.pdf" \
  -F "pages=1,2,3" \
  -o result.docx
```

## Деплой на Railway

1. Залей эту папку в свой GitHub-репозиторий:
   ```bash
   git init
   git add .
   git commit -m "pdf-to-word backend"
   git branch -M main
   git remote add origin https://github.com/<твой-юзернейм>/pdf-to-word-backend.git
   git push -u origin main
   ```
2. На [railway.app](https://railway.app) → New Project → Deploy from GitHub repo → выбери этот репозиторий.
3. Railway сам подхватит `nixpacks.toml` (ставит `poppler-utils`) и `railway.json` (команда старта `npm start`).
4. В настройках проекта → Variables добавь:
   - `ANTHROPIC_API_KEY` — твой ключ с [console.anthropic.com](https://console.anthropic.com)
   - (опционально) `ANTHROPIC_MODEL` — по умолчанию `claude-sonnet-5`
5. После деплоя Railway выдаст публичный URL вида `https://<project>.up.railway.app` — это и есть `API_BASE_URL` для расширения (см. `config.js` в папке расширения).

## Ограничения текущей версии (честно)

- Изображения со страницы (логотипы, фото) сейчас не переносятся в .docx — переносится только текст и таблицы, которые распознаёт модель. Перенос картинок как бинарных данных — следующий шаг (нужно доставать embedded images из PDF отдельно, не через vision).
- Каждая страница — отдельный вызов Claude API, то есть стоимость и время растут линейно с числом страниц. Для книг на сотни страниц обязательно используйте параметр `pages`, чтобы не гонять всё через API.
- Нет очереди/прогресса по SSE — это синхронный запрос, ответ приходит целиком, когда все страницы обработаны. Для больших PDF стоит добавить стриминг прогресса (see TODO in server.js).
- Нет лимитов/биллинга/авторизации — это база для прототипа, не production-ready сервис.
