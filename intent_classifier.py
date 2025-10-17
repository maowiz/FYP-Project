# intent_classifier.py
import os
import joblib
import numpy as np
import random
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import SGDClassifier
from sklearn.pipeline import Pipeline
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report, confusion_matrix
import warnings
warnings.filterwarnings('ignore')

class IntentClassifier:
    def __init__(self, command_handler_instance, retrain=False):
        self.model_filename = 'intent_classifier.pkl'
        self.command_handler = command_handler_instance
        self.pipeline = None
        self.label_map = {0: 'command', 1: 'general_query'}
        
        if os.path.exists(self.model_filename) and not retrain:
            print("Loading existing intent classifier model...")
            self.load_model()
        else:
            print("Training new intent classifier model...")
            self.train_model()

    def _prepare_training_data(self):
        """
        Generates comprehensive training data from commands and general queries.
        """
        X_train = []
        y_train = []
        
        # ============= PART 1: COMMANDS (Label 0) =============
        print("Generating command examples...")
        
        # Direct commands
        commands = list(self.command_handler.COMMANDS.keys())
        for cmd in commands:
            X_train.append(cmd)
            y_train.append(0)
            
            # Add natural variations
            variations = self._generate_command_variations(cmd)
            for var in variations:
                X_train.append(var)
                y_train.append(0)
        
        # Command synonyms (Corrected logic)
        for command, synonyms in self.command_handler.COMMAND_SYNONYMS.items():
            for synonym in synonyms:
                # Skip synonyms with placeholders like {n} as they are not good for training
                if '{' in synonym and '}' in synonym:
                    continue
                
                X_train.append(synonym)
                y_train.append(0)
                
                # Add some variations for important synonyms
                if random.random() > 0.5:  # 50% chance to add variations
                    X_train.append(f"please {synonym}")
                    y_train.append(0)
        
        print(f"Generated {len([y for y in y_train if y == 0])} command examples")
        
        # ============= PART 2: GENERAL QUERIES (Label 1) =============
        print("Generating general query examples...")
        general_queries = self._generate_comprehensive_general_queries()
        X_train.extend(general_queries)
        y_train.extend([1] * len(general_queries))
        
        print(f"Generated {len(general_queries)} general query examples")
        
        # ============= PART 3: AMBIGUOUS EXAMPLES =============
        ambiguous = self._generate_ambiguous_examples()
        for amb, label in ambiguous:
            X_train.append(amb)
            y_train.append(label)
        
        print(f"Total training examples: {len(X_train)}")
        print(f"Command examples: {y_train.count(0)} ({y_train.count(0)/len(y_train)*100:.1f}%)")
        print(f"General query examples: {y_train.count(1)} ({y_train.count(1)/len(y_train)*100:.1f}%)")
        
        return X_train, y_train

    def _generate_command_variations(self, command):
        """Generate natural variations of a command."""
        variations = []
        
        # Common prefixes
        prefixes = [
            "please", "can you", "could you", "would you",
            "I want to", "I need to", "let's", "just",
            "quickly", "hey", "okay", "go ahead and",
            "help me", "I'd like to", "will you"
        ]
        
        # Common suffixes
        suffixes = [
            "please", "now", "for me", "right now",
            "immediately", "quickly", "asap"
        ]
        
        # Select random variations (but not too many to avoid overfitting)
        num_variations = min(5, len(prefixes))
        selected_prefixes = random.sample(prefixes, num_variations)
        
        for prefix in selected_prefixes:
            variations.append(f"{prefix} {command}")
        
        # Add some suffix variations
        for suffix in random.sample(suffixes, min(2, len(suffixes))):
            variations.append(f"{command} {suffix}")
        
        return variations

    def _generate_comprehensive_general_queries(self):
        """Generate diverse general queries that should be handled by LLM."""
        queries = []
        
        # What questions
        what_questions = [
            "what is the meaning of life",
            "what is quantum computing",
            "what's the weather like today",
            "what do you think about artificial intelligence",
            "what is the capital of Japan",
            "what time is it in London",
            "what's the difference between ML and AI",
            "what should I cook for dinner",
            "what's the best programming language",
            "what is photosynthesis",
            "what's happening in the news today",
            "what is blockchain technology",
            "what's your favorite color",
            "what is the theory of relativity",
            "what causes rain",
            "what's the population of China",
            "what is machine learning",
            "what are black holes",
            "what's the meaning of this word",
            "what happened in World War 2",
            "what is climate change",
            "what's the recipe for chocolate cake",
            "what are the symptoms of flu",
            "what is cryptocurrency",
            "what's the distance to the moon",
            "what makes people happy",
            "what's the stock market",
            "what are neural networks",
            "what is democracy",
            "what's the speed of light"
        ]
        
        # How questions
        how_questions = [
            "how does a computer work",
            "how to learn programming",
            "how do you make coffee",
            "how does photosynthesis work",
            "how can I improve my memory",
            "how to solve a rubik's cube",
            "how does the internet work",
            "how to be more productive",
            "how do airplanes fly",
            "how to cook pasta perfectly",
            "how does machine learning work",
            "how to write a resume",
            "how can I learn faster",
            "how to meditate",
            "how does a car engine work",
            "how to save money effectively",
            "how do vaccines work",
            "how to start a business",
            "how does electricity work",
            "how to be happy in life",
            "how to write better code",
            "how does the brain work",
            "how to lose weight healthily",
            "how to manage stress",
            "how do solar panels work",
            "how to speak confidently",
            "how to manage time better",
            "how to study effectively",
            "how to make friends",
            "how to invest money"
        ]
        
        # Can you questions
        can_you_questions = [
            "can you tell me a joke",
            "can you explain quantum physics",
            "can you help me understand calculus",
            "can you write a poem about love",
            "can you suggest a movie to watch",
            "can you recommend a book",
            "can you explain how AI works",
            "can you help me with my homework",
            "can you tell me a story",
            "can you give me advice about relationships",
            "can you explain cryptocurrency",
            "can you help me learn Spanish",
            "can you suggest a workout routine",
            "can you explain the stock market",
            "can you write code for a calculator",
            "can you help me debug this code",
            "can you explain machine learning algorithms",
            "can you suggest healthy recipes",
            "can you help me plan a trip",
            "can you explain climate change",
            "can you write a haiku",
            "can you solve this math problem",
            "can you recommend music",
            "can you help me write an email",
            "can you explain blockchain",
            "can you translate this text",
            "can you summarize this article",
            "can you check my grammar",
            "can you teach me something new",
            "can you analyze this data"
        ]
        
        # Tell me requests
        tell_me_questions = [
            "tell me about the solar system",
            "tell me a fun fact",
            "tell me about yourself",
            "tell me something interesting",
            "tell me about ancient Egypt",
            "tell me how computers evolved",
            "tell me about the ocean",
            "tell me about space exploration",
            "tell me the history of the internet",
            "tell me about dinosaurs",
            "tell me about the human brain",
            "tell me something I don't know",
            "tell me about renewable energy",
            "tell me about different cultures",
            "tell me about the future of technology",
            "tell me about famous scientists",
            "tell me about the universe",
            "tell me about healthy eating",
            "tell me about meditation benefits",
            "tell me about global warming",
            "tell me about artificial intelligence ethics",
            "tell me about the stock market history",
            "tell me about different religions",
            "tell me about mental health",
            "tell me about the environment",
            "tell me about World War 2",
            "tell me about the Renaissance",
            "tell me about quantum physics",
            "tell me about evolution",
            "tell me about cryptocurrency"
        ]
        
        # Explain requests
        explain_questions = [
            "explain the theory of evolution",
            "explain how the internet works",
            "explain quantum computing to me",
            "explain what DNA is",
            "explain the water cycle",
            "explain how electricity is generated",
            "explain the concept of time",
            "explain how memory works",
            "explain artificial intelligence",
            "explain the greenhouse effect",
            "explain how vaccines work",
            "explain the big bang theory",
            "explain cryptocurrency mining",
            "explain how search engines work",
            "explain machine learning algorithms",
            "explain the stock market crash",
            "explain how GPS works",
            "explain nuclear fusion",
            "explain how WiFi works",
            "explain the immune system",
            "explain climate change causes",
            "explain how batteries work",
            "explain social media algorithms",
            "explain how credit cards work",
            "explain the speed of light",
            "explain it like I'm 5",
            "explain in simple terms",
            "explain with examples",
            "explain step by step",
            "explain the basics"
        ]
        
        # Why questions
        why_questions = [
            "why is the sky blue",
            "why do we dream",
            "why is water important",
            "why do leaves change color",
            "why is exercise important",
            "why do we need sleep",
            "why is the ocean salty",
            "why do we have seasons",
            "why is biodiversity important",
            "why do we age",
            "why is the sun hot",
            "why do we have emotions",
            "why is coding important",
            "why do birds migrate",
            "why is math important",
            "why do we have different languages",
            "why is reading beneficial",
            "why do we need vitamins",
            "why is history important",
            "why do economies crash",
            "why is privacy important online",
            "why do we have laws",
            "why is renewable energy important",
            "why do we have time zones",
            "why is mental health important",
            "why do people lie",
            "why do we fall in love",
            "why do bad things happen",
            "why should I save money",
            "why learn history"
        ]
        
        # Who/Where/When questions
        who_where_when_questions = [
            "who invented the computer",
            "who is the richest person",
            "who discovered electricity",
            "who wrote Romeo and Juliet",
            "who invented the internet",
            "where is the Great Wall of China",
            "where can I learn programming",
            "where is the deepest ocean",
            "where do hurricanes form",
            "where is Silicon Valley",
            "when was the internet invented",
            "when is the best time to exercise",
            "when did World War 2 end",
            "when will we colonize Mars",
            "when was the first computer made",
            "who was Albert Einstein",
            "where is Mount Everest",
            "when did dinosaurs exist",
            "who invented Bitcoin",
            "where is the Amazon rainforest"
        ]
        
        # Conversational and personal
        conversational = [
            "hello how are you",
            "good morning",
            "how's your day going",
            "what do you think",
            "that's interesting",
            "I don't understand",
            "could you elaborate",
            "what else can you tell me",
            "thanks for the help",
            "you're very helpful",
            "I have a question",
            "let me think about that",
            "that makes sense",
            "I see what you mean",
            "could you repeat that",
            "what's your opinion",
            "do you agree",
            "I'm not sure about that",
            "that's a good point",
            "let's talk about something else",
            "nice to meet you",
            "introduce yourself",
            "who are you",
            "what are you",
            "what's your name"
        ]
        
        # Creative and help requests
        creative_help = [
            "write a Python function to sort a list",
            "generate a business plan outline",
            "create a workout schedule for me",
            "help me understand this concept",
            "I'm feeling stressed what should I do",
            "give me tips for public speaking",
            "suggest some healthy breakfast ideas",
            "help me write a cover letter",
            "I want to learn a new skill",
            "draft an email to my boss",
            "give me study tips",
            "help me plan my day",
            "I need gift ideas for my friend",
            "write a short story about space",
            "help me solve this problem",
            "give me motivation to work",
            "suggest ways to save money",
            "help me make a decision",
            "I'm bored what should I do",
            "create a meal plan for the week",
            "give me interview tips",
            "help me be more productive",
            "suggest a hobby to try",
            "write a thank you note",
            "I need advice about my career",
            "help me organize my tasks",
            "give me ideas for a project",
            "summarize this article for me",
            "translate this to Spanish",
            "write a poem about nature"
        ]
        
        # Opinion and thought questions
        opinion_questions = [
            "what's your opinion on climate change",
            "do you think AI will replace humans",
            "what do you think about the future",
            "in your opinion what's most important",
            "what's your take on this",
            "do you agree that technology is good",
            "what would you say about politics",
            "from your perspective",
            "what's your view on education",
            "what's your stance on privacy"
        ]
        
        # Combine all categories
        queries.extend(what_questions)
        queries.extend(how_questions)
        queries.extend(can_you_questions)
        queries.extend(tell_me_questions)
        queries.extend(explain_questions)
        queries.extend(why_questions)
        queries.extend(who_where_when_questions)
        queries.extend(conversational)
        queries.extend(creative_help)
        queries.extend(opinion_questions)
        
        return queries

    def _generate_ambiguous_examples(self):
        """Generate examples that could be either commands or queries."""
        ambiguous = [
            ("open discussion about files", 1),  # Likely query
            ("create a story about folders", 1),  # Likely query
            ("search for the best restaurants", 0),  # Could be search command
            ("tell me the current time please", 0),  # Could be tell time command
            ("volume of a sphere formula", 1),  # Likely query about math
            ("how to minimize stress in life", 1),  # Likely life advice
            ("maximize your potential", 1),  # Likely motivational
            ("close the discussion please", 1),  # Likely conversation
            ("switch topics please", 1),  # Likely conversation
            ("what's the current brightness level", 0),  # Could be brightness command
            ("play some relaxing music", 0),  # Could be play command
            ("check my email please", 0),  # Could be email command
            ("find information about Python", 0),  # Could be search command
            ("google how to cook pasta", 0),  # Search command
            ("youtube tutorial on programming", 0),  # YouTube command
        ]
        return ambiguous

    def train_model(self):
        """Train the classification model with comprehensive dataset."""
        X_train, y_train = self._prepare_training_data()
        
        # Split for validation
        X_train_split, X_val, y_train_split, y_val = train_test_split(
            X_train, y_train, test_size=0.2, random_state=42, stratify=y_train
        )
        
        # Create pipeline with optimized parameters
        self.pipeline = Pipeline([
            ('tfidf', TfidfVectorizer(
                max_features=2000,
                ngram_range=(1, 3),
                min_df=2,
                max_df=0.95,
                use_idf=True,
                smooth_idf=True,
                sublinear_tf=True
            )),
            ('classifier', SGDClassifier(
                loss='log_loss',
                penalty='l2',
                alpha=0.0001,
                random_state=42,
                max_iter=200,
                tol=1e-3,
                n_jobs=-1,
                class_weight='balanced'
            ))
        ])
        
        print("\nTraining classifier...")
        self.pipeline.fit(X_train_split, y_train_split)
        
        # Evaluate on validation set
        print("\nValidation Results:")
        y_pred = self.pipeline.predict(X_val)
        print(classification_report(y_val, y_pred, 
                                   target_names=['command', 'general_query']))
        
        # Retrain on full dataset
        print("\nRetraining on full dataset...")
        self.pipeline.fit(X_train, y_train)
        
        self.save_model()
        print("Classifier training complete and model saved.")

    def classify(self, text: str, confidence_threshold: float = 0.6):
        """
        Classify text with confidence threshold.
        Returns (intent, confidence, should_use_llm)
        """
        if not self.pipeline:
            raise RuntimeError("Classifier model is not loaded or trained.")
        
        # Get probabilities
        probabilities = self.pipeline.predict_proba([text.lower()])[0]
        predicted_class_index = np.argmax(probabilities)
        
        intent = self.label_map[predicted_class_index]
        confidence = probabilities[predicted_class_index]
        
        # Determine if we should use LLM based on confidence
        if confidence < confidence_threshold:
            # Low confidence - might be ambiguous
            should_use_llm = True
        elif intent == 'general_query':
            should_use_llm = True
        else:
            should_use_llm = False
        
        return intent, confidence, should_use_llm

    def save_model(self):
        """Save the trained model."""
        joblib.dump(self.pipeline, self.model_filename)
        print(f"Model saved to {self.model_filename}")

    def load_model(self):
        """Load a previously trained model."""
        self.pipeline = joblib.load(self.model_filename)
        print(f"Model loaded from {self.model_filename}")

    def test_classifier(self):
        """Test the classifier with various examples."""
        test_examples = [
            # Clear commands
            ("create folder", "command"),
            ("please open my computer", "command"),
            ("increase volume to 80", "command"),
            ("take a screenshot now", "command"),
            ("minimize all windows please", "command"),
            ("play music on youtube", "command"),
            
            # Clear queries
            ("what is the meaning of life", "general_query"),
            ("how does quantum computing work", "general_query"),
            ("can you write a poem", "general_query"),
            ("explain machine learning to me", "general_query"),
            ("tell me about ancient history", "general_query"),
            
            # Ambiguous
            ("search for information", "?"),
            ("tell me the time", "?"),
            ("open discussion", "?"),
        ]
        
        print("\n" + "="*60)
        print("CLASSIFIER TEST RESULTS")
        print("="*60)
        
        for text, expected in test_examples:
            intent, confidence, use_llm = self.classify(text)
            status = "✓" if intent == expected or expected == "?" else "✗"
            print(f"{status} '{text}'")
            print(f"   -> Intent: {intent}, Confidence: {confidence:.2%}, Use LLM: {use_llm}")
            print()

# Main execution block
if __name__ == '__main__':
    print("="*60)
    print("INTENT CLASSIFIER TRAINING")
    print("="*60)
    
    # Import dependencies
    from command_handler import CommandHandler
    from file_management import FileManager
    from os_management import OSManagement
    
    # Initialize components
    os_manager = OSManagement(None)
    file_manager = FileManager(None, os_manager, None)
    command_handler_instance = CommandHandler(file_manager, os_manager)
    
    # Train classifier (force retrain)
    classifier = IntentClassifier(command_handler_instance, retrain=True)
    
    # Test the classifier
    classifier.test_classifier()
    
    # Additional custom tests
    print("\n" + "="*60)
    print("CUSTOM TEST EXAMPLES")
    print("="*60)
    
    custom_tests = [
        "can you please create a folder for my new project",
        "what is your opinion on artificial intelligence",
        "open youtube and play some music",
        "how do I become a better programmer",
        "maximize the volume please",
        "write a story about a dragon",
        "check my internet speed",
        "why is the sky blue",
        "take a photo now",
        "hello how are you doing today",
    ]
    
    for test in custom_tests:
        intent, confidence, use_llm = classifier.classify(test)
        print(f"'{test}'")
        print(f"   -> Intent: {intent}, Confidence: {confidence:.2%}, Use LLM: {use_llm}")
        print()