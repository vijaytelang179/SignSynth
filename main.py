import sys
import os
import requests
import webbrowser
import subprocess
import tempfile
import time
import json
import queue
import string
import threading

from PyQt5.QtWidgets import QApplication, QMessageBox
from loading_screen import LoadingScreen

from panda3d.core import loadPrcFileData, Filename

APP_VERSION = "v1.1.0"
GITHUB_REPO = "Suja2004/ASR"

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS

    plugins_dir = os.path.join(base_path, "plugins")
    plugins_dir = Filename.fromOsSpecific(plugins_dir).toOsSpecific()
    loadPrcFileData("", f"plugin-path {plugins_dir}")

    model_path = Filename.fromOsSpecific(base_path).toOsSpecific()
    loadPrcFileData("", f"model-path {model_path}")

    loadPrcFileData("", "load-display pandagl")
    loadPrcFileData("", "aux-display tinydisplay")
    loadPrcFileData("", "window-title SignSynth")
else:
    base_path = os.path.abspath(os.path.dirname(__file__))

    model_path = Filename.fromOsSpecific(base_path).toOsSpecific()
    loadPrcFileData("", f"model-path {model_path}")

if getattr(sys, 'frozen', False):
    dll_path = os.path.join(sys._MEIPASS, "vosk")
    if os.path.exists(dll_path):
        os.add_dll_directory(dll_path)

from panda3d.core import LVecBase3f, DirectionalLight, AmbientLight, TextNode, WindowProperties

import pyaudio
import win32com.client
from direct.gui.DirectButton import DirectButton
from direct.gui.DirectFrame import DirectFrame
from direct.gui.OnscreenText import OnscreenText
from direct.interval.IntervalGlobal import Sequence
from direct.interval.LerpInterval import LerpPosInterval, LerpHprInterval
from direct.showbase.ShowBase import ShowBase
from direct.task import Task
from nltk.corpus import stopwords, wordnet
from nltk.tokenize import word_tokenize
from vosk import Model, KaldiRecognizer
from nltk import pos_tag, WordNetLemmatizer

import nltk

try:
    nltk.data.find('tokenizers/punkt_tab')
    nltk.data.find('tokenizers/punkt')
    nltk.data.find('corpora/stopwords')
    nltk.data.find('corpora/wordnet')
    nltk.data.find('taggers/averaged_perceptron_tagger')
    nltk.data.find('corpora/omw-1.4')
except LookupError:
    nltk.download('punkt_tab', quiet=True)
    nltk.download('punkt', quiet=True)
    nltk.download('stopwords', quiet=True)
    nltk.download('wordnet', quiet=True)
    nltk.download('averaged_perceptron_tagger', quiet=True)
    nltk.download('omw-1.4', quiet=True)


class ContinuousSpeechGloss:
    """
    Continuously recognizes speech, converts it to sign language gloss,
    and passes results to a callback or queue.
    """

    def __init__(self, callback=None):
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)

        model_path = os.path.join(base_path, "vosk-model-small-en-us-0.15")

        self.lemmatizer = WordNetLemmatizer()

        self.stop_words = set(stopwords.words('english')) - {
            'i', 'you', 'we', 'he', 'she', 'they', 'me', 'my', 'your', 'our', 'his', 'her', 'their'
        }

        self.gloss_map = {
            "i": "ME",
            "you": "YOU",
            "we": "US",
            "he": "HE",
            "she": "SHE",
            "they": "THEY",
            "am": "",
            "is": "",
            "'s": "",
            "n't": "",
            "'re": "",
            "'ve": "",
            "are": "",
            "was": "",
            "were": "",
            "going": "GO",
            "go": "GO",
            "had": "HAVE",
            "don't": "NOT",
            "not": "NOT",
            "no": "NOT",
            "won't": "NOT WILL",
            "store": "STORE",
            "because": "WHY",
            "milk": "MILK",
            "to": "",
            "the": "",
            "a": "",
            "an": "",
            "and": "PLUS",
            "but": "BUT",
            "this": "THIS",
            "that": "THAT",
            "there": "THERE",
            "here": "HERE",
            "what": "WHAT",
            "who": "WHO",
            "where": "WHERE",
            "when": "WHEN",
            "why": "WHY",
            "hello": "HI",
            "talk": "SPEAK",
            "learn": "LEARN",
            "try": "TRY",
            "coached": "COACH",
            "habits": "HABIT",
            "millions": "MILLION",
            "skills": "SKILL",
            "think": "OVERTHINKING"
        }

        self.model_path = model_path
        self.callback = callback
        self.running = False
        self.thread = None
        self.results = queue.Queue()

    def convert_to_sign_gloss(self, text):
        words = [w for w in word_tokenize(text.lower()) if w not in string.punctuation]
        pos_tags = pos_tag(words)

        def get_wordnet_pos(tag):
            if tag.startswith('J'):
                return wordnet.ADJ
            elif tag.startswith('V'):
                return wordnet.VERB
            elif tag.startswith('N'):
                return wordnet.NOUN
            elif tag.startswith('R'):
                return wordnet.ADV
            else:
                return wordnet.NOUN

        lemmatized_words = [self.lemmatizer.lemmatize(w, get_wordnet_pos(t)) for w, t in pos_tags]

        gloss_sequence = []
        seen_pronouns = set()
        for word in lemmatized_words:
            if word in self.stop_words and word not in self.gloss_map:
                continue
            gloss_word = self.gloss_map.get(word, word.upper()).strip()
            if not gloss_word:
                continue
            if gloss_word in {"ME", "YOU", "HE", "SHE", "US", "THEY"}:
                if gloss_word in seen_pronouns:
                    continue
                seen_pronouns.add(gloss_word)
            gloss_sequence.append(gloss_word)

        return " ".join(gloss_sequence)

    def start(self):
        """Start continuous speech recognition in a background thread"""
        if self.running:
            return False

        self.running = True
        self.thread = threading.Thread(target=self._listen_continuously)
        self.thread.daemon = True
        self.thread.start()
        return True

    def stop(self):
        """Stop background speech recognition"""
        self.running = False
        if self.thread:
            self.thread.join(timeout=1.0)
        return True

    def get_latest_result(self):
        """Retrieve latest recognition result from internal queue (if no callback used)"""
        if not self.results.empty():
            return self.results.get()
        return None

    def _listen_continuously(self):
        """
        Constantly listens for speech, performs recognition and gloss conversion.
        Sends results via callback or internal queue.
        """
        try:
            model = Model(self.model_path)
            recognizer = KaldiRecognizer(model, 16000)

            mic = pyaudio.PyAudio()
            stream = mic.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=16000,
                input=True,
                frames_per_buffer=8192
            )
            stream.start_stream()

            print("Continuous speech recognition started...")

            while self.running:
                data = stream.read(4096, exception_on_overflow=False)

                if recognizer.AcceptWaveform(data):
                    result = json.loads(recognizer.Result())
                    text = result.get("text", "").strip()
                    if text:
                        gloss = self.convert_to_sign_gloss(text)
                        if self.callback:
                            self.callback(text, gloss)
                        else:
                            self.results.put((text, gloss))

                time.sleep(0.01)

            stream.stop_stream()
            stream.close()
            mic.terminate()
            print("Continuous speech recognition stopped.")

        except Exception as e:
            error_msg = f"Error in speech recognition: {str(e)}"
            print(error_msg)
            if self.callback:
                self.callback(error_msg, "")
            else:
                self.results.put((error_msg, ""))
            self.running = False


class SignLanguageApp(ShowBase):
    """
    Main application class: integrates 3D model, sign pose animation, UI,
    speech recognition, and optional media control for sign language display.
    """

    def __init__(self):

        ShowBase.__init__(self)

        self.loadModels()
        self.setupLights()
        self.setupSkybox()

        try:
            self.current_pose = "default"
            self.gesture_data = self.loadAllPoseData()
            self.loadSignPoses(self.current_pose)
            self.expanded_sequence = []
            self.pose_index = 0
            self.is_animating = False
        except Exception as e:
            print(f"Could not load pose data: {e}")

        self.media_control_active = False
        self.play_interval = 5
        self.pause_interval = 5
        self.last_media_action_time = 0
        self.media_state = "paused"

        self.speech_recognition_active = False
        self.speech_processor = None

        self.is_animating = False
        self.signing_complete = True

        self.setup_ui()

        self.start_speech_recognition()

        self.setup_media_control()

    def open_app_window(self):
        """
        Manually opens the main Panda3D window and runs all
        window-dependent setup code.
        """

        if self.openDefaultWindow():
            print("Successfully opened Panda3D window.")

            props = WindowProperties()
            props.setTitle("SignSynth")

            icon_path = self.get_resource_path("SignSynth.ico")

            if os.path.exists(icon_path):
                panda_icon_path = Filename.fromOsSpecific(icon_path)
                props.setIconFilename(panda_icon_path)
            else:
                print(f"Warning: Icon file not found at {icon_path}")

            self.win.requestProperties(props)
            self.disableMouse()
            self.camera.setPos(0, -15, 3.25)
            self.camera.lookAt(0, 0, 0)

            self.title = OnscreenText(
                text="SignSynth",
                style=1, fg=(1, 1, 1, 1), pos=(0, 0.9),
                scale=0.1, mayChange=True
            )

        else:
            print("Error: Failed to open Panda3D window.")

    def setup_ui(self):
        """
        Create a consolidated, on-screen user interface panel.
        """
        COLOR_ACTIVE = (0.9, 0.3, 0.3, 1)
        COLOR_INACTIVE = (0.3, 0.6, 0.9, 1)
        FRAME_COLOR = (0.1, 0.1, 0.1, 0)
        TEXT_COLOR = (1, 1, 1, 1)

        self.ui_frame = DirectFrame(
            frameColor=FRAME_COLOR,
            frameSize=(-1.3, 1.3, -0.25, 0.25),
            pos=(0.4, 0, -0.7)
        )

        self.top_bar_frame = DirectFrame(
            frameColor=FRAME_COLOR,
            frameSize=(-1.3, 1.3, -0.25, 0.25),
            pos=(0.1, 0, 0.5)
        )

        self.recognized_text_label = OnscreenText(
            parent=self.ui_frame, text="Current Sign:", pos=(-1.2, 0.1), scale=0.06,
            fg=TEXT_COLOR, align=TextNode.ALeft, mayChange=False
        )

        self.recognized_text_node = OnscreenText(
            parent=self.ui_frame, text="...", pos=(-0.65, 0.1), scale=0.06,
            fg=TEXT_COLOR, align=TextNode.ALeft, mayChange=True
        )

        self.gloss_text_label = OnscreenText(
            parent=self.ui_frame, text="Signing (Gloss):", pos=(-1.2, -0.1), scale=0.06,
            fg=TEXT_COLOR, align=TextNode.ALeft, mayChange=False
        )

        self.gloss_text_node = OnscreenText(
            parent=self.ui_frame, text="Ready to listen.", pos=(-0.65, -0.1), scale=0.06,
            fg=TEXT_COLOR, align=TextNode.ALeft, wordwrap=20, mayChange=True
        )

        self.speech_toggle_button = DirectButton(
            parent=self.top_bar_frame,
            text="Speech",
            text_scale=0.05,
            frameSize=(-0.2, 0.2, -0.08, 0.08),
            command=self.toggle_speech_recognition,
            pos=(0.6, 0, 0.1),
            frameColor=COLOR_ACTIVE,
            relief='raised',
            borderWidth=(0.01, 0.01)
        )

        self.media_toggle_button = DirectButton(
            parent=self.top_bar_frame,
            text="Media Control",
            text_scale=0.05,
            frameSize=(-0.2, 0.2, -0.08, 0.08),
            command=self.toggle_media_control,
            pos=(0.6, 0, -0.1),
            frameColor=COLOR_INACTIVE,
            relief='raised',
            borderWidth=(0.01, 0.01)
        )

    def loadModels(self):
        """Load 3D character model, arms, and attach to scene graph."""
        try:

            body_path = self.get_panda_model_path('character/body.bam')
            print(f"Loading body from: {body_path}")
            self.torso = self.loader.loadModel(body_path)
            self.torso.reparentTo(self.render)
            self.torso.setPos(0, 0, -1.5)
            self.torso.setScale(0.7)
            self.torso.setHpr(0, 0, 0)

            rarm_path = self.get_panda_model_path('character/RArm.bam')
            print(f"Loading right arm from: {rarm_path}")
            self.rarm = self.loader.loadModel(rarm_path)
            self.rarm.reparentTo(self.torso)

            larm_path = self.get_panda_model_path('character/LArm.bam')
            print(f"Loading left arm from: {larm_path}")
            self.larm = self.loader.loadModel(larm_path)
            self.larm.reparentTo(self.torso)

            self.setup_arm_details()
            print("All models loaded successfully")

        except Exception as e:
            print(f"Error loading models: {e}")
            import traceback
            traceback.print_exc()
            raise

    def setup_arm_details(self):
        """
        Store references to all finger segment nodes for both arms
        so poses can be efficiently applied.
        """
        self.rthumb1 = self.rarm.find("**/t1")
        self.rthumb2 = self.rarm.find("**/t2")
        self.rindex1 = self.rarm.find("**/i1")
        self.rindex2 = self.rarm.find("**/i2")
        self.rindex3 = self.rarm.find("**/i3")
        self.rmiddle1 = self.rarm.find("**/m1")
        self.rmiddle2 = self.rarm.find("**/m2")
        self.rmiddle3 = self.rarm.find("**/m3")
        self.rring1 = self.rarm.find("**/r1")
        self.rring2 = self.rarm.find("**/r2")
        self.rring3 = self.rarm.find("**/r3")
        self.rpinky1 = self.rarm.find("**/p1")
        self.rpinky2 = self.rarm.find("**/p2")
        self.rpinky3 = self.rarm.find("**/p3")

        self.lthumb1 = self.larm.find("**/t1")
        self.lthumb2 = self.larm.find("**/t2")
        self.lindex1 = self.larm.find("**/i1")
        self.lindex2 = self.larm.find("**/i2")
        self.lindex3 = self.larm.find("**/i3")
        self.lmiddle1 = self.larm.find("**/m1")
        self.lmiddle2 = self.larm.find("**/m2")
        self.lmiddle3 = self.larm.find("**/m3")
        self.lring1 = self.larm.find("**/r1")
        self.lring2 = self.larm.find("**/r2")
        self.lring3 = self.larm.find("**/r3")
        self.lpinky1 = self.larm.find("**/p1")
        self.lpinky2 = self.larm.find("**/p2")
        self.lpinky3 = self.larm.find("**/p3")

    def setupLights(self):
        """Create a main directional light and ambient light for 3D scene."""
        mainLight = DirectionalLight('main light')
        mainLight.setShadowCaster(True)
        mainLightNodePath = self.render.attachNewNode(mainLight)
        mainLightNodePath.setHpr(0, -40, 0)

        self.render.setLight(mainLightNodePath)

        ambientLight = AmbientLight('ambient light')
        ambientLight.setColor((0.2, 0.2, 0.2, 1))
        ambientLightNodePath = self.render.attachNewNode(ambientLight)
        self.render.setLight(ambientLightNodePath)
        self.render.setShaderAuto()

    def setupSkybox(self):
        """Loads a skybox model if available, otherwise prints an error."""
        try:
            skybox = self.loader.loadModel('skybox/skybox.bam')
            skybox.setScale(50)
            skybox.setBin('background', 1)
            skybox.setDepthWrite(0)
            skybox.setLightOff()
            skybox.reparentTo(self.render)
        except Exception as e:
            print(f"Could not load skybox: {e}")

    def get_resource_path(self, relative_path):
        """Get absolute path to resource, works for dev and PyInstaller"""
        if getattr(sys, 'frozen', False):

            base_path = sys._MEIPASS
        else:

            base_path = os.path.abspath(os.path.dirname(__file__))

        return os.path.join(base_path, relative_path)

    def get_panda_model_path(self, relative_path):
        """
        Get Panda3D-compatible model path.
        NOTE: This now just returns the relative path, as the
        model-path is set correctly at startup for both dev and exe.
        """
        return relative_path

    def loadAllPoseData(self):
        """Load all sign pose definitions from local JSON file."""
        pose_file = self.get_resource_path("sign_poses.json")

        try:
            with open(pose_file, "r") as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Error: Could not find sign_poses.json at {pose_file}")
            raise
        except json.JSONDecodeError as e:
            print(f"Error: Invalid JSON in sign_poses.json: {e}")
            raise

    def loadSignPoses(self, name):
        """
        Apply pose data for a given sign or letter to hand and finger models.
        """
        poses = self.gesture_data.get(name)
        pose = poses[0] if isinstance(poses, list) else poses
        if not pose:
            return
        l = pose["leftHand"]
        r = pose["rightHand"]

        self.larm.setPos(LVecBase3f(*l["pos"]))
        self.larm.setHpr(LVecBase3f(*l["hpr"]))
        self.rarm.setPos(LVecBase3f(*r["pos"]))
        self.rarm.setHpr(LVecBase3f(*r["hpr"]))

        def applyFingerPose(finger_parts, data):
            for part, pose_data in zip(finger_parts, data):
                part.setPos(*pose_data["pos"])
                part.setHpr(*pose_data["hpr"])

        if "fingers" in l:
            f = l["fingers"]
            if "thumb" in f:
                applyFingerPose([self.lthumb1, self.lthumb2], f["thumb"])
            if "index" in f:
                applyFingerPose([self.lindex1, self.lindex2,
                                 self.lindex3], f["index"])
            if "middle" in f:
                applyFingerPose([self.lmiddle1, self.lmiddle2,
                                 self.lmiddle3], f["middle"])
            if "ring" in f:
                applyFingerPose(
                    [self.lring1, self.lring2, self.lring3], f["ring"])
            if "pinky" in f:
                applyFingerPose([self.lpinky1, self.lpinky2,
                                 self.lpinky3], f["pinky"])

        if "fingers" in r:
            f = r["fingers"]
            if "thumb" in f:
                applyFingerPose([self.rthumb1, self.rthumb2], f["thumb"])
            if "index" in f:
                applyFingerPose([self.rindex1, self.rindex2,
                                 self.rindex3], f["index"])
            if "middle" in f:
                applyFingerPose([self.rmiddle1, self.rmiddle2,
                                 self.rmiddle3], f["middle"])
            if "ring" in f:
                applyFingerPose(
                    [self.rring1, self.rring2, self.rring3], f["ring"])
            if "pinky" in f:
                applyFingerPose([self.rpinky1, self.rpinky2,
                                 self.rpinky3], f["pinky"])

    def expandPoseSequence(self, sequence):
        """
        For each recognized word, use that sign if available; otherwise,
        expand to fingerspelling (letter signs).
        """
        result = []
        for word in sequence:
            if word.lower() in self.gesture_data:
                result.append(word.lower())
            else:
                for letter in word.lower():
                    if letter in self.gesture_data:
                        result.append(letter)
        return result

    def start_animation(self, text):
        """
        Start animating signs for a recognized phrase or sentence.
        """
        self.stopAnimation()
        self.current_text = text.strip()
        words = self.current_text.split()
        self.expanded_sequence = self.expandPoseSequence(words)

        if not self.expanded_sequence:
            self.gloss_text_node.setText("No valid signs found in text")
            self.signing_complete = True
            return

        self.gloss_text_node.setText(f"Signing: {self.current_text}")
        self.pose_index = 0
        self.is_animating = True
        self.signing_complete = False

        if self.media_control_active and self.media_state == "playing":
            self.pause_media()
        self.taskMgr.add(self.animateNextPose, "SignAnimation")

    def stopAnimation(self):

        if self.is_animating:
            self.taskMgr.remove("SignAnimation")
            self.is_animating = False

            if hasattr(self, 'current_left_seq') and self.current_left_seq:
                self.current_left_seq.finish()
                self.current_left_seq = None
            if hasattr(self, 'current_right_seq') and self.current_right_seq:
                self.current_right_seq.finish()
                self.current_right_seq = None

    def slideArms(self):
        """
        Provide a simple sliding motion animation for fingerspelling transitions.
        """
        slide_distance = 0.5
        time = 0.2
        sequence = Sequence(
            LerpPosInterval(self.larm, time, self.larm.getPos()),
            LerpPosInterval(self.rarm, time, self.rarm.getPos() +
                            LVecBase3f(-slide_distance, 0, 0)),
            LerpPosInterval(self.larm, time, self.larm.getPos()),
            LerpPosInterval(self.rarm, time, self.rarm.getPos())
        )
        sequence.start()

    def animateNextPose(self, task):
        """
        Animates each sign or letter in the expanded sequence, applying relevant poses.
        """
        if self.pose_index >= len(self.expanded_sequence):

            if hasattr(self, 'current_left_seq') and self.current_left_seq and self.current_left_seq.isPlaying():
                return task.again
            if hasattr(self, 'current_right_seq') and self.current_right_seq and self.current_right_seq.isPlaying():
                return task.again

            self.loadSignPoses("default")
            self.pose_index = 0
            self.is_animating = False
            self.gloss_text_node.setText("Animation Complete")
            self.current_pose = ""

            self.signing_complete = True

            self.current_left_seq = None
            self.current_right_seq = None

            if self.media_control_active and self.media_state == "paused":
                self.resume_media()
            return Task.done

        pose_name = self.expanded_sequence[self.pose_index]

        if self.current_pose == pose_name and len(pose_name) == 1:
            self.slideArms()
            self.pose_index += 1
            return task.again

        self.current_pose = pose_name
        poses = self.gesture_data.get(pose_name)
        if not poses:
            self.pose_index += 1
            return task.again

        left_sequence = []
        right_sequence = []
        time = 0.005

        def addFingerLerp(hand_data, finger_map, sequence_list):
            if "fingers" not in hand_data:
                return
            for name, parts in finger_map.items():
                if name in hand_data["fingers"]:
                    for part, pose_data in zip(parts, hand_data["fingers"][name]):
                        sequence_list.append(LerpPosInterval(
                            part, 0.01, LVecBase3f(*pose_data["pos"])))
                        sequence_list.append(LerpHprInterval(
                            part, 0.01, LVecBase3f(*pose_data["hpr"])))

        def addHandAndFingers(pose):
            l = pose["leftHand"]
            r = pose["rightHand"]

            left_sequence.extend([
                LerpPosInterval(self.larm, time, LVecBase3f(*l["pos"])),
                LerpHprInterval(self.larm, time, LVecBase3f(*l["hpr"]))
            ])
            addFingerLerp(l, {
                "thumb": [self.lthumb1, self.lthumb2],
                "index": [self.lindex1, self.lindex2, self.lindex3],
                "middle": [self.lmiddle1, self.lmiddle2, self.lmiddle3],
                "ring": [self.lring1, self.lring2, self.lring3],
                "pinky": [self.lpinky1, self.lpinky2, self.lpinky3]
            }, left_sequence)

            right_sequence.extend([
                LerpPosInterval(self.rarm, time, LVecBase3f(*r["pos"])),
                LerpHprInterval(self.rarm, time, LVecBase3f(*r["hpr"]))
            ])
            addFingerLerp(r, {
                "thumb": [self.rthumb1, self.rthumb2],
                "index": [self.rindex1, self.rindex2, self.rindex3],
                "middle": [self.rmiddle1, self.rmiddle2, self.rmiddle3],
                "ring": [self.rring1, self.rring2, self.rring3],
                "pinky": [self.rpinky1, self.rpinky2, self.rpinky3]
            }, right_sequence)

        if isinstance(poses, list):
            for pose in poses:
                addHandAndFingers(pose)
        else:
            addHandAndFingers(poses)

        self.current_left_seq = None
        self.current_right_seq = None

        if left_sequence:
            self.current_left_seq = Sequence(*left_sequence)
            self.current_left_seq.start()

        if right_sequence:
            self.current_right_seq = Sequence(*right_sequence)
            self.current_right_seq.start()

        self.gloss_text_node.setText(f"Signing: {self.current_text}")
        self.recognized_text_node.setText(f"{pose_name.upper()}")

        task.delayTime = 1.5
        self.pose_index += 1
        return task.again

    def setup_media_control(self):
        """
        Set up background task for media control (play/pause).
        Takes no effect until activated.
        """
        self.taskMgr.add(self.media_control_task, "MediaControlTask")

    def toggle_media_control(self):
        """
        Enable or disable media control mode. Updates UI and resets timing.
        """
        try:
            self.media_control_active = not self.media_control_active

            if self.media_control_active:
                self.media_toggle_button["text"] = "Media Control"
                self.media_toggle_button["frameColor"] = (0.9, 0.3, 0.3, 1)
                self.gloss_text_node.setText(
                    "Media control starting (switch to media tab)")
                self.last_media_action_time = time.time()
                self.media_state = "starting"
                print(
                    "Media control starting - switch to your media tab within 3 seconds!")
            else:
                self.media_toggle_button["text"] = "Media Control"
                self.media_toggle_button["frameColor"] = (0.3, 0.6, 0.9, 1)
                self.gloss_text_node.setText(
                    "Speech recognition active, media control inactive")
                self.media_state = "paused"
                print("Media control stopped")

        except Exception as e:
            print(f"Error toggling media control: {str(e)}")
            self.gloss_text_node.setText(f"Error: {str(e)}")

    def media_control_task(self, task):
        """
        Background task for alternating media play/pause based on timing.
        Only active if media_control_active flag is set.
        """
        if not self.media_control_active:
            return Task.cont
        if not self.signing_complete:
            return Task.cont

        current_time = time.time()
        elapsed = current_time - self.last_media_action_time

        if self.media_state == "starting" and elapsed >= 3:

            self.last_media_action_time = current_time
            self.media_state = "playing"
            self.gloss_text_node.setText("Media playing")

        elif self.media_state == "playing" and elapsed >= self.play_interval:
            self.pause_media()

        return Task.cont

    def pause_media(self):
        """
        Simulate keyboard press to pause media and update internal state/UI.
        """
        self.simulate_space_press()
        self.last_media_action_time = time.time()
        self.media_state = "paused"
        self.gloss_text_node.setText("Media paused")

    def resume_media(self):
        """
        Simulate keyboard press to resume media and update internal state/UI.
        """
        self.simulate_space_press()
        self.last_media_action_time = time.time()
        self.media_state = "playing"
        self.gloss_text_node.setText("Media playing")

    def simulate_space_press(self):
        """
        Programmatically sends a space key press to control external media playback.
        Handles cross-platform differences.
        """
        if sys.platform == 'win32':
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys(" ", 0)
            except ImportError:
                print(
                    "Could not import win32com.client - media control may not work properly")
        else:
            try:
                import pyautogui
                pyautogui.press('space')
            except ImportError:
                print("Could not import pyautogui - media control may not work properly")

    def start_speech_recognition(self):
        try:
            if not self.speech_processor:
                self.speech_processor = ContinuousSpeechGloss(callback=self.handle_speech_result)
            if self.speech_processor.start():
                self.speech_recognition_active = True
                self.speech_toggle_button["text"] = "Speech"
                self.speech_toggle_button["frameColor"] = (0.9, 0.3, 0.3, 1)
                self.gloss_text_node.setText("Speech active. Ready to listen.")
            else:
                self.gloss_text_node.setText("Error: Speech failed to start")
        except Exception as e:
            self.gloss_text_node.setText(f"Error: {str(e)}")

    def toggle_speech_recognition(self):
        if not self.speech_recognition_active:
            self.start_speech_recognition()
        else:
            try:
                if self.speech_processor and self.speech_processor.stop():
                    self.speech_recognition_active = False
                    self.speech_toggle_button["text"] = "Speech"
                    self.speech_toggle_button["frameColor"] = (0.3, 0.6, 0.9, 1)
                    self.gloss_text_node.setText("Speech inactive.")
                else:
                    self.gloss_text_node.setText("Error: Failed to stop speech")
            except Exception as e:
                self.gloss_text_node.setText(f"Error: {str(e)}")

    def handle_speech_result(self, text, gloss):
        """
        Called whenever new speech is recognized.
        Updates UI and triggers sign animation (unless an animation is already running).
        """
        if text and gloss and not self.is_animating:
            self.recognized_text_node.setText(text)
            self.gloss_text_node.setText(gloss)
            self.start_animation(gloss)


def check_for_updates(loader):
    """
    Checks GitHub for the latest release and prompts user to update.
    Returns True to continue loading, False to quit.
    """
    loader.update_progress("Checking for updates...", "Connecting to GitHub...")
    loader.update()

    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        response = requests.get(api_url, timeout=10)
        response.raise_for_status()

        latest_release = response.json()
        latest_version = latest_release.get("tag_name", "")

        def version_tuple(v):
            """Convert version string to tuple for comparison"""
            try:
                return tuple(map(int, v.lstrip('v').split('.')))
            except:
                return (0, 0, 0)

        try:
            current_ver = version_tuple(APP_VERSION)
            latest_ver = version_tuple(latest_version)
            update_available = latest_ver > current_ver
        except (ValueError, AttributeError):
            print(f"Could not parse versions: current={APP_VERSION}, latest={latest_version}")
            update_available = False

        if update_available:
            print(f"Update found: {latest_version} (current: {APP_VERSION})")

            installer_url = None
            for asset in latest_release.get("assets", []):
                if asset.get("name", "").endswith(".exe"):
                    installer_url = asset.get("browser_download_url")
                    break

            if not installer_url:
                print("Update found, but no .exe installer available.")
                loader.update_progress("Update available", "No installer found, continuing...")
                loader.update()
                time.sleep(1)
                return True

            choice = {'value': None}

            def on_yes():
                choice['value'] = True
                loader.quit()

            def on_no():
                choice['value'] = False
                loader.quit()

            loader.show_update_prompt(
                f"A new version ({latest_version}) is available!",
                on_yes,
                on_no
            )

            loader.mainloop()

            if choice['value'] == True:

                try:
                    loader.update_progress("Downloading update...", "Starting download...")
                    loader.update()

                    temp_dir = tempfile.gettempdir()
                    installer_name = os.path.basename(installer_url)
                    installer_path = os.path.join(temp_dir, installer_name)

                    with requests.get(installer_url, stream=True, timeout=30) as r:
                        r.raise_for_status()
                        total_size = int(r.headers.get('content-length', 0))
                        downloaded_size = 0

                        with open(installer_path, 'wb') as f:
                            for chunk in r.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)
                                    downloaded_size += len(chunk)

                                    progress_mb = f"{(downloaded_size / (1024 * 1024)):.1f} MB"
                                    if total_size > 0:
                                        progress_mb += f" / {(total_size / (1024 * 1024)):.1f} MB"
                                        progress_pct = int((downloaded_size / total_size) * 100)
                                        loader.set_progress(progress_pct)

                                    loader.update_progress("Downloading update...", progress_mb)
                                    loader.update()

                    loader.update_progress("Update downloaded", "Launching installer...")
                    loader.update()
                    time.sleep(1)

                    if sys.platform == 'win32':
                        subprocess.Popen([installer_path, '/SILENT'])
                    else:
                        subprocess.Popen([installer_path])

                    return False

                except Exception as e:
                    print(f"Failed to download or run updater: {e}")
                    loader.hide_update_prompt()
                    loader.update_progress("Download failed", "Opening browser instead...")
                    loader.update()

                    import tkinter.messagebox as messagebox
                    messagebox.showwarning(
                        "Update Failed",
                        f"Could not automatically download the update.\n\nError: {str(e)}\n\nOpening the browser instead."
                    )
                    webbrowser.open(installer_url)
                    return False

            else:

                loader.hide_update_prompt()
                loader.update_progress("Continuing with current version...", "")
                loader.update()
                time.sleep(0.5)
                return True
        else:

            print(f"App is up-to-date (version {APP_VERSION})")
            loader.update_progress("Application is up-to-date", "")
            loader.update()
            time.sleep(0.5)
            return True

    except requests.exceptions.Timeout:
        print("Update check timed out")
        loader.update_progress("Update check timed out", "Continuing offline...")
        loader.update()
        time.sleep(1)
        return True

    except requests.exceptions.RequestException as e:
        print(f"Could not check for updates: {e}")
        loader.update_progress("Could not check for updates", "Continuing offline...")
        loader.update()
        time.sleep(1)
        return True

    except Exception as e:
        print(f"Unexpected error during update check: {e}")
        loader.update_progress("Update check failed", "Continuing...")
        loader.update()
        time.sleep(1)
        return True


if __name__ == "__main__":

    loading = LoadingScreen(version=APP_VERSION)

    loading.set_steps([
        "Checking for updates",
        "Initializing 3D engine",
        "Loading 3D models",
        "Initializing audio",
        "Finalizing UI"
    ])

    loading.center()
    loading.show()
    loading.update()

    should_continue = check_for_updates(loading)

    if not should_continue:
        loading.close()
        sys.exit()

    loadPrcFileData("", "window-type none")

    loading.update_progress("Initializing 3D engine...", "Starting Panda3D...")
    loading.update()

    try:
        panda_app = SignLanguageApp()
    except Exception as e:
        loading.close()
        import tkinter.messagebox as messagebox

        messagebox.showerror("Fatal Error", f"Failed to initialize Panda3D: {e}")
        sys.exit(1)

    loading.update_progress("Loading 3D models...", "Character, arms, and skybox")
    loading.update()

    loading.update_progress("Initializing audio...", "Starting Vosk speech engine")
    loading.update()

    loading.update_progress("Finalizing UI...", "Preparing user interface")
    loading.update()


    def on_loading_finished():
        print("Loading complete. Starting Panda3D event loop.")
        panda_app.open_app_window()
        panda_app.run()


    loading.finished_connect(on_loading_finished)
    loading.complete()

    loading.mainloop()