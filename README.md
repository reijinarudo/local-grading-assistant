# ReggieGPT

ReggieGPT is a private AI assistant that runs on one computer you own and is reachable from other machines on your home or office network. It combines an open-weight language model, a browser chat interface, and a set of Python grading tools, so an instructor can draft feedback on student work without sending that work to a cloud AI service.

## Why it exists

Instructors who use AI to help grade face a privacy problem: student submissions are protected records, and schools are tightening rules on sending them to third-party services. ReggieGPT keeps every file on local hardware. The model, the chat interface, and the grading scripts all run on the same machine, with no cloud fallback.

The project has three layers:

1. **Model server.** llama.cpp's `llama-server` loads a GGUF model and serves an OpenAI-compatible API on the local machine only.
2. **Chat interface.** Open WebUI provides the browser chat window, file uploads, per-model system prompts, and filters. Other computers on the LAN open it in a browser.
3. **Grading tools.** `make_packet.py` reads a student's Word workbook, SPSS data file, and SPSS output file, recomputes the statistics, checks the files for common problems, and writes a plain-text evidence packet. An Open WebUI filter runs that script automatically when files are attached to a chat, so the model grades from verified numbers.

The model reads and writes about material placed in front of it. It is an unreliable source of facts from its own recall, which is why the grading tools put every number the model needs into the packet.

## Demo

[Add a screenshot of Open WebUI with a generated evidence packet, or a short clip, to `docs/`.]

A sample evidence packet built from the fictional data in `examples/` begins like this:

```
EVIDENCE PACKET  |  Example Module: Describing a Data Set  |  Student: Sample Student
...
SECTION 2. AUTOMATED FLAGS (observations to verify against the text below)
  - Report word count: 36 (limit 150). Within limit.
  - Report names a specific major: business, nursing
...
  StudyHours (N = 60)
    mean 8.77  MATCH
    SD 4.014  MATCH
```

## Hardware

- Reference machine: Mac mini with an Apple M4 chip and 24 GB of unified memory.
- Any machine that runs llama.cpp will work. The limit is memory: the model file, plus its KV cache at your chosen context length, plus the operating system, must fit with headroom to spare.
- In the reference setup, a 27-billion-parameter model at 4-bit quantization (about 14 GB on disk) with a 32K context left about 27 percent of memory free. It generated about 6.5 tokens per second and read prompts at about 60 tokens per second.

## Software

| Component | Version tested | Purpose |
| --- | --- | --- |
| macOS | 26.5 | Reference OS. The grading tools also run on Linux and Windows. |
| llama.cpp (`llama-server`) | Homebrew build, 2026 | Model server |
| A GGUF model | Qwen3.8-27B, Unsloth UD-IQ4_XS quantization | Language model |
| Python | 3.11 for Open WebUI, 3.10 or newer for the grading tools | |
| uv | current | Creates the Python environments |
| Open WebUI | 0.11.3 | Chat interface |
| python-docx, pyreadstat, pandas, pillow | see `requirements.txt` | Grading tools |

## Installation

These steps assume macOS with Homebrew. On Linux, install llama.cpp from your package manager or build it from source; the rest is the same.

1. Clone the repository and enter it.

   ```
   git clone https://github.com/reijinarudo/reggiegpt.git
   cd reggiegpt
   ```

2. Install llama.cpp and uv.

   ```
   brew install llama.cpp uv
   ```

3. Download a GGUF model that fits your memory and note its path. Hugging Face hosts many; check each model card for its license.

4. Install Open WebUI in its own environment. Open WebUI's documentation lists the Python versions it supports; 3.11 was used here.

   ```
   uv venv --python 3.11 ~/open-webui/.venv
   uv pip install --python ~/open-webui/.venv/bin/python3 open-webui
   ```

5. Install the grading libraries into Open WebUI's environment, so the filter can import them.

   ```
   uv pip install --python ~/open-webui/.venv/bin/python3 -r requirements.txt
   ```

6. Create your settings file.

   ```
   cp .env.example .env
   ```

   Open `.env` and set at least `MODEL_PATH`.

## Configuration

All settings live in `.env`. Each line in `.env.example` has a comment saying what it does. The ones you are most likely to change:

- `MODEL_PATH` and `MODEL_ALIAS`: the model file and the name it appears under in Open WebUI.
- `MODEL_CTX`: context length. Longer documents need more, and memory use rises with it.
- `CHAT_TEMPLATE_KWARGS`: model-specific template switches. For Qwen3-family models, `'{"enable_thinking":false}'` turns off the reasoning trace, which cut the time to first word from minutes to about a second in testing.
- `WEBUI_HOST` and `WEBUI_PORT`: where the browser interface listens. Open WebUI's own default port is 8080; set the port you want explicitly.

The grading tools read a course config file. Copy `grader/course_config.example.json` to `grader/course_config.json` and replace the fictional values with your assignment's box names, variables, codebook, and the true values of your data set. `course_config.json` is listed in `.gitignore` because it holds an answer key.

## Usage

### Start the services

```
./server/start-model.sh      # first terminal
./server/start-webui.sh      # second terminal
```

Open `http://<server-name>.local:3000` from any machine on your LAN, or `http://localhost:3000` on the server itself. Create the admin account when prompted. Accounts are stored in a local database on the server.

To start both services at login on macOS, use the templates in `server/launchd/`. Each file explains its three install steps.

### Two Open WebUI settings that matter for speed

- **Admin Panel > Settings > Models > your model > Capabilities: turn off Builtin Tools.** With it on, Open WebUI 0.11 sent about 6,000 tokens of tool definitions with every message, which delayed each reply by over a minute on this hardware.
- **Settings > Interface: turn off automatic title, follow-up, and tag generation.** Each one is an extra model call.

### Try the grading tools on the fictional sample

```
python3 examples/make_sample_data.py
python3 grader/make_packet.py examples/sample_submission \
    --config grader/course_config.example.json \
    --source-sav examples/sample_course_data.sav \
    --out packets
```

The packet appears at `packets/Sample_Student_packet.txt`, and the pasted charts appear as PNG files in `packets/Sample_Student_charts/`. The sample has no SPSS output file, so the packet reports it as missing.

### Grade inside Open WebUI

1. Copy `grader/` to the folder named in `GRADER_DIR`, with your `course_config.json`, the blank assignment workbook saved as `blank_workbook.docx`, and the course data file saved as `course_data.sav`. The last two are optional; they let the script separate student text from instructions and detect an unchanged data file.
2. Fill in `grader/grading_brief.example.txt` for your assignment and save it as `grading_brief.txt`.
3. In Open WebUI, go to Admin Panel > Functions, add a new function, paste the contents of `grader/openwebui_filter_evidence_packet.py`, save, and switch it on. Use the gear icon to check the folder paths.
4. Admin Panel > Settings > Models > your grading model: paste the brief into System Prompt, tick the "Evidence Packet" filter, and save.
5. Open a new chat on that model, attach one student's files, and type `grade this`.

Use one chat per student, so no student's work carries into the next grade. Before relying on the output, test it against work you have already graded; `docs/evaluating-a-local-grader.md` describes how.

## Project layout

```
reggiegpt/
  README.md
  LICENSE                                  MIT, for this project's own code
  .env.example                             every setting, with placeholder values
  requirements.txt                         grading tool dependencies
  server/
    start-model.sh                         starts llama-server from .env
    start-webui.sh                         starts Open WebUI from .env
    launchd/*.plist.example                macOS auto-start templates
  grader/
    make_packet.py                         builds the evidence packet
    openwebui_filter_evidence_packet.py    Open WebUI filter that runs make_packet.py
    course_config.example.json             fictional course config
    grading_brief.example.txt              system prompt template
  examples/
    make_sample_data.py                    generates a fictional submission
  docs/
    evaluating-a-local-grader.md           how to test a model before trusting it
```

## Credits

ReggieGPT is built on the following open-source projects and models. This repository contains none of their source code; it installs and configures them.

| Project | Author | Link | License |
| --- | --- | --- | --- |
| llama.cpp | Georgi Gerganov and contributors (ggml-org) | https://github.com/ggml-org/llama.cpp | MIT |
| Open WebUI | Timothy Jaeryang Baek and contributors | https://github.com/open-webui/open-webui | Open WebUI License (BSD-3-Clause with a branding clause) |
| Qwen models | Qwen team, Alibaba Cloud | https://huggingface.co/Qwen | See each model card |
| GGUF quantizations | Unsloth AI | https://huggingface.co/unsloth | See each model card |
| python-docx | Steve Canny and contributors | https://github.com/python-openxml/python-docx | MIT |
| pyreadstat | Otto Fajardo and contributors (Roche) | https://github.com/Roche/pyreadstat | Apache 2.0 |
| Pillow | Jeffrey A. Clark and contributors | https://github.com/python-pillow/Pillow | MIT-CMU |
| uv | Astral | https://github.com/astral-sh/uv | MIT or Apache 2.0 |

Open WebUI's license asks that its branding stay visible in deployments; read it before customizing the interface.

## Modifications and original work

Everything below is Reginald Finley's work on top of the upstream projects.

- **System design.** Chose and combined the three layers for a privacy requirement: student data never leaves the machine, the model listens on the local machine only, and the interface is reachable on the LAN only.
- **Model server configuration.** Selected a model and quantization to fit a 24 GB machine, measured memory headroom and speed, and settled on one slot, an 8-bit KV cache, and template-level thinking control after testing showed the default reasoning mode took minutes per reply.
- **Open WebUI tuning.** Found the Builtin Tools payload that was adding over a minute to every reply, by reading the raw request body in llama-server's verbose log, and documented the settings that remove it.
- **Evidence packet builder (`make_packet.py`).** Extracts each answer box and its pasted charts from a Word workbook; recomputes descriptive statistics from the student's SPSS file, with SPSS-matching skewness and Tukey-hinge quartiles alongside the linear method; checks value and variable labels; lists the SPSS procedures in the output file; scans the workbook for hidden text, tracked changes, and text boxes; and flags common report errors. All assignment-specific facts load from a JSON config.
- **Open WebUI filter.** Intercepts attached student files, builds the packet, and replaces the raw uploads with the packet before the model sees the message.
- **Grading brief template and evaluation method.** A system prompt structure written for a small model, and a pass/fail test for deciding whether a local model is fit to draft grades.
- **Fictional sample data** so anyone can run the tools without real student work.

The code in `grader/` and `examples/` was written with Claude (Anthropic) as a coding assistant, under Reginald Finley's direction, with his specification, testing, and review.

## License

The code and documentation in this repository are released under the MIT License; see `LICENSE`. The upstream projects and models listed under Credits keep their own licenses.

## Contact

Reginald Finley, PhD. contact@drreginaldfinley.com. Portfolio: https://drreginaldfinley.com
