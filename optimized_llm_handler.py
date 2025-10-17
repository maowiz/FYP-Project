# optimized_llm_handler.py (Improved Version)

import os
import sys
import multiprocessing
from llama_cpp import Llama

class OptimizedLLMHandler:

    def __init__(self, model_path=None):
        """
        Initializes the LLM handler with robust error checking and dynamic
        CPU thread allocation.
        """
        # 0. If no model path is provided, use the default. This makes the
        #    handler more robust against being passed `None`.
        if model_path is None:
            model_path = r"C:\Users\SC\Desktop\FYP-Project-main\FYP-Project-main\qwen2.5-0.5b-instruct-q4_k_m.gguf"

        # 1. Check if the model file actually exists before trying to load it.
        if not os.path.exists(model_path):
            print(f"FATAL ERROR: Model file not found at '{os.path.abspath(model_path)}'")
            print("Please ensure the model has been downloaded and is in the project's main directory.")
            sys.exit(1) # Exit the program if the model is missing.

        try:
            print("Loading LLM... This may take a moment.")
            
            # 2. Automatically detect CPU cores for optimal thread settings.
            # This makes the code more portable to different computers.
            cpu_cores = multiprocessing.cpu_count()
            
            self.llm = Llama(
                model_path=model_path,
                n_ctx=512,
                n_batch=8,
                # Leave at least one core free for the OS and other processes.
                n_threads=max(1, cpu_cores - 1),
                n_threads_batch=max(1, (cpu_cores - 1) // 2),
                use_mmap=True,
                verbose=False,
                seed=-1 # Use a random seed for varied conversational responses.
            )
            
            # 3. Use the improved prompt template with examples.
            # This helps the LLM better understand its task.
            self.function_template = """<|im_start|>system
You are Aura, a helpful voice assistant. Interpret the user's request and respond in one of two ways:
1.  If the request is a system command, respond ONLY with the command in this format: CMD: exact_command_name
2.  For any other request (greetings, questions, conversation), provide a natural, brief, and helpful answer.

Here are the available system commands: {functions}

Examples:
User: "create a new folder for me"
Assistant: CMD: create folder

User: "what is the weather in Khanewal?"
Assistant: The weather in Khanewal is clear with a temperature of around 34 degrees Celsius.

User: "hello"
Assistant: Hello! How can I help you today?

User: "what is your name?"
Assistant: I am Aura, your voice assistant.

User: "explain quantum computing"
Assistant: Quantum computing uses principles from quantum mechanics to solve complex problems much faster than classical computers.<|im_end|>
<|im_start|>user
{user_input}<|im_end|>
<|im_start|>assistant
"""
            # 4. A dedicated prompt for essay generation.
            self.essay_template = """<|im_start|>system
You are a helpful writing assistant. Your task is to write a well-structured, multi-paragraph essay on the topic provided by the user. Respond ONLY with the essay content. Do not add any extra commentary or conversational text.
<|im_end|>
<|im_start|>user
{user_input}<|im_end|>
<|im_start|>assistant
"""
            # 4. A dedicated prompt for translation tasks.
            self.translation_template = """<|im_start|>system
You are a multilingual translation expert. Your task is to translate the user's phrase into the specified language accurately. Respond ONLY with the translated text. Do not add any extra commentary or explanations.
<|im_end|>
<|im_start|>user
Translate the phrase "{phrase}" to {language}.<|im_end|>
<|im_start|>assistant
"""
            # 5. A dedicated prompt for summarization.
            self.summarization_template = """<|im_start|>system
You are an expert summarization assistant. Your task is to provide a concise, easy-to-read summary of the text provided by the user. Respond ONLY with the summarized text.
<|im_end|>
<|im_start|>user
Summarize the following text:
{text_to_summarize}<|im_end|>
<|im_start|>assistant
"""
            print("LLM loaded successfully.")

        except Exception as e:
            print(f"FATAL ERROR: Failed to load the LLM. Error: {e}")
            # This will re-raise the error to stop the program if the LLM can't be loaded.
            raise

    def count_tokens(self, text: str) -> int:
        """Counts the number of tokens in a given text string."""
        # Add a space to handle single-word inputs correctly
        return len(self.llm.tokenize(f" {text}".encode("utf-8")))

    def process_fast(self, text, available_functions=None):
        """
        Processes the user's text using the LLM, streaming the response.
        Yields the generated text and whether it's a command.
        """
        # The available_functions list is now passed into the improved prompt
        function_list_str = ", ".join(available_functions) if available_functions else "none"
        
        prompt = self.function_template.format(
            functions=function_list_str,
            user_input=text
        )
        
        response_text = ""
        is_command = False
        
        # Stream the response token by token
        for token in self.llm(
            prompt,
            max_tokens=80, # Slightly increased for better conversational answers
            temperature=0.2, # Low temperature for more predictable responses
            stream=True,
            stop=["<|im_end|>", "\n", "User:"] # Stop generating if it hallucinates a new turn
        ):
            chunk = token['choices'][0]['text']
            response_text += chunk
            
            # Early detection of a command
            if "CMD:" in response_text and not is_command:
                is_command = True

        # Yield the final, complete response
        yield response_text.strip(), is_command

    def generate_essay(self, prompt: str):
        """
        Generates essay content using a dedicated essay prompt.
        This is a direct, non-streaming method.
        """
        full_prompt = self.essay_template.format(user_input=prompt)

        response = self.llm(
            full_prompt,
            max_tokens=350, # Allow for a longer, more detailed essay
            temperature=0.5, # A bit more creative for writing
            stream=False, # No streaming needed for this internal call
            stop=["<|im_end|>", "\n\n\n", "User:"]
        )

        if response and 'choices' in response and response['choices']:
            essay_text = response['choices'][0]['text'].strip()
            # Clean up potential LLM artifacts
            if essay_text.startswith("Assistant:"):
                essay_text = essay_text[len("Assistant:"):].strip()
            if essay_text.startswith("Here is an essay"):
                essay_text = essay_text[essay_text.find('\n'):].strip()
            return essay_text
        
        return "I'm sorry, I couldn't generate an essay on that topic."

    def generate_summary(self, text_to_summarize: str):
        """
        Generates a summary of the provided text.
        """
        full_prompt = self.summarization_template.format(text_to_summarize=text_to_summarize)

        response = self.llm(
            full_prompt,
            max_tokens=150, # Summaries should be concise
            temperature=0.3,
            stream=False,
            stop=["<|im_end|>", "\n\n"]
        )

        if response and 'choices' in response and response['choices']:
            summary_text = response['choices'][0]['text'].strip()
            # Clean up potential LLM artifacts
            if summary_text.startswith("Assistant:"):
                summary_text = summary_text[len("Assistant:"):].strip()
            return summary_text
        
        return "I'm sorry, I couldn't generate a summary for that text."

    def process_translation(self, phrase: str, language: str):
        """
        Processes a translation request using a dedicated prompt.
        Yields the translated text.
        """
        prompt = self.translation_template.format(
            phrase=phrase,
            language=language
        )

        response_text = ""

        # Stream the response token by token
        for token in self.llm(
            prompt,
            max_tokens=100, # Allow for longer translated phrases
            temperature=0.1, # Very low temperature for accurate translation
            stream=True,
            stop=["<|im_end|>", "\n", "User:"]
        ):
            chunk = token['choices'][0]['text']
            response_text += chunk

        # Yield the final, complete response. The second value (is_command) is False.
        yield response_text.strip(), False
