import json
import queue
import string
import sys
import threading
import time

import nltk
import pyaudio
import win32com.client
from direct.gui.DirectButton import DirectButton
from direct.gui.DirectFrame import DirectFrame
from direct.gui.OnscreenText import OnscreenText
from direct.interval.IntervalGlobal import Sequence
from direct.interval.LerpInterval import LerpPosInterval, LerpHprInterval
from direct.showbase.ShowBase import ShowBase
from direct.task import Task
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from panda3d.core import LVecBase3f, DirectionalLight, AmbientLight
from vosk import Model, KaldiRecognizer


class ContinuousSpeechGloss:
    """
    Continuously recognizes speech, converts it to sign language gloss,
    and passes results to a callback or queue.
    """
    def __init__(self, model_path="vosk-model-small-en-us-0.15", callback=None):
        # Ensure required nltk resources are available
        try:
            nltk.data.find('tokenizers/punkt')
            nltk.data.find('corpora/stopwords')
        except LookupError:
            nltk.download('punkt', quiet=True)
            nltk.download('stopwords', quiet=True)

        # Set up stop words and pronouns required for glossing
        self.stop_words = set(stopwords.words('english')) - {
            'i', 'you', 'we', 'he', 'she', 'they', 'me', 'my', 'your', 'our', 'his', 'her', 'their'
        }

        # Mapping from words/phrases to gloss
        self.gloss_map = {
            "i": "ME", "you": "YOU", "we": "US", "he": "HE", "she": "SHE", "they": "THEY",
            "am": "", "is": "", "are": "", "was": "", "were": "",
            "going": "GO", "go": "GO", "want": "WANT", "have": "HAVE", "had": "HAVE",
            "don't": "NOT", "not": "NOT", "no": "NOT", "won't": "NOT WILL",
            "store": "STORE", "because": "WHY", "milk": "MILK", "to": "",
            "the": "", "a": "", "an": "", "and": "PLUS", "but": "BUT",
            "this": "THIS", "that": "THAT", "there": "THERE", "here": "HERE",
            "what": "WHAT", "who": "WHO", "where": "WHERE", "when": "WHEN", "why": "WHY", "how": "HOW",
            "need": "NEED", "can": "CAN", "will": "WILL", "should": "SHOULD", "must": "MUST",
            "good": "GOOD", "bad": "BAD", "happy": "HAPPY", "sad": "SAD",
            "yes": "YES", "okay": "OK", "like": "LIKE", "help": "HELP"
        }

        self.model_path = model_path
        self.callback = callback
        self.running = False
        self.thread = None
        self.results = queue.Queue()

    def convert_to_sign_gloss(self, text):
        """Convert normal text to sign language gloss notation"""
        words = word_tokenize(text.lower())
        words = [word for word in words if word not in string.punctuation]
        filtered = [word for word in words if word not in self.stop_words or word.lower() in self.gloss_map]

        gloss_sequence = []
        for word in filtered:
            gloss_word = self.gloss_map.get(word.lower(), word.upper())
            if gloss_word:
                gloss_sequence.append(gloss_word)

        gloss_string = " ".join(gloss_sequence)
        return gloss_string

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

        # Camera and visual setup
        self.disableMouse()
        self.camera.setPos(0, -15, 3.25)
        self.camera.lookAt(0, 0, 0)

        # Title display
        self.title = OnscreenText(
            text="SignSynth",
            style=1, fg=(1, 1, 1, 1), pos=(0, 0.9),
            scale=0.1, mayChange=True
        )

        # Load all 3D models and scene elements
        self.loadModels()
        self.setupLights()
        self.setupSkybox()

        # Load pose definitions for all signs
        try:
            self.current_pose = "default"
            self.gesture_data = self.loadAllPoseData()
            self.loadSignPoses(self.current_pose)
            self.expanded_sequence = []
            self.pose_index = 0
            self.is_animating = False
        except Exception as e:
            print(f"Could not load pose data: {e}")

        # Media control state variables
        self.media_control_active = False
        self.play_interval = 5
        self.pause_interval = 5
        self.last_media_action_time = 0
        self.media_state = "paused"

        # Speech recognition state
        self.speech_recognition_active = False
        self.speech_processor = None

        # Animation flags
        self.signing_complete = True  # Signals whether signing is in progress

        # Set up the user interface
        self.setup_ui()

        # Start the speech recognition processor on launch
        self.start_speech_recognition()

        # Prepare (but do not activate) media control functionality
        self.setup_media_control()

    def setup_ui(self):
        """
        Create on-screen user interface elements for:
          - Status display
          - Control buttons for speech and media control
        """
        self.status_text = OnscreenText(
            text="Ready for speech input. Media control inactive.",
            pos=(0, -0.55),
            scale=0.06,
            fg=(1, 1, 1, 1),
            wordwrap=30,
            mayChange=True
        )

        self.control_frame = DirectFrame(
            frameColor=(0.2, 0.2, 0.2, 0.7),
            frameSize=(-0.7, 0.7, -0.15, 0.15),
            pos=(0, 0, -0.8)
        )

        self.speech_toggle_button = DirectButton(
            parent=self.control_frame,
            text="Stop Speech Recognition",
            text_scale=0.04,
            frameSize=(-0.3, 0.3, -0.06, 0.06),
            command=self.toggle_speech_recognition,
            pos=(-0.35, 0, 0),
            frameColor=(0.9, 0.3, 0.3, 1)
        )

        self.media_toggle_button = DirectButton(
            parent=self.control_frame,
            text="Start Media Control",
            text_scale=0.04,
            frameSize=(-0.3, 0.3, -0.06, 0.06),
            command=self.toggle_media_control,
            pos=(0.35, 0, 0),
            frameColor=(0.3, 0.6, 0.9, 1)
        )

    def loadModels(self):
        """Load 3D character model, arms, and attach to scene graph."""
        self.torso = loader.loadModel('character/torso.glb')
        self.torso.setPos(0, 0, -1.5)
        self.torso.reparentTo(render)
        self.torso.setScale(0.7)

        # Load left and right arms as children
        self.rarm = loader.loadModel('character/RArmX.glb')
        self.rarm.reparentTo(self.torso)
        self.larm = loader.loadModel('character/LArmX.glb')
        self.larm.reparentTo(self.torso)
        # Set up references to finger parts for each arm
        self.setup_arm_details()

    def setup_arm_details(self):
        """
        Store references to all finger segment nodes for both arms
        so poses can be efficiently applied.
        """
        # Right arm fingers
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

        # Left arm fingers
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
        mainLightNodePath = render.attachNewNode(mainLight)
        mainLightNodePath.setHpr(0, -70, 0)
        render.setLight(mainLightNodePath)

        ambientLight = AmbientLight('ambient light')
        ambientLight.setColor((0.2, 0.2, 0.2, 1))
        ambientLightNodePath = render.attachNewNode(ambientLight)
        render.setLight(ambientLightNodePath)
        render.setShaderAuto()

    def setupSkybox(self):
        """Loads a skybox model if available, otherwise prints an error."""
        try:
            skybox = loader.loadModel('skybox/skybox.egg')
            skybox.setScale(50)
            skybox.setBin('background', 1)
            skybox.setDepthWrite(0)
            skybox.setLightOff()
            skybox.reparentTo(render)
        except Exception as e:
            print(f"Could not load skybox: {e}")

    def loadAllPoseData(self):
        """Load all sign pose definitions from local JSON file."""
        with open("sign_poses.json", "r") as f:
            return json.load(f)

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
        self.larm.setPos(*l["pos"])
        self.larm.setHpr(*l["hpr"])
        self.rarm.setPos(*r["pos"])
        self.rarm.setHpr(*r["hpr"])

        def applyFingerPose(finger_parts, data):
            for part, pose_data in zip(finger_parts, data):
                part.setPos(*pose_data["pos"])
                part.setHpr(*pose_data["hpr"])

        if "fingers" in l:
            f = l["fingers"]
            if "thumb" in f:
                applyFingerPose([self.lthumb1, self.lthumb2], f["thumb"])
            if "index" in f:
                applyFingerPose([self.lindex1, self.lindex2, self.lindex3], f["index"])
            if "middle" in f:
                applyFingerPose([self.lmiddle1, self.lmiddle2, self.lmiddle3], f["middle"])
            if "ring" in f:
                applyFingerPose([self.lring1, self.lring2, self.lring3], f["ring"])
            if "pinky" in f:
                applyFingerPose([self.lpinky1, self.lpinky2, self.lpinky3], f["pinky"])

        if "fingers" in r:
            f = r["fingers"]
            if "thumb" in f:
                applyFingerPose([self.rthumb1, self.rthumb2], f["thumb"])
            if "index" in f:
                applyFingerPose([self.rindex1, self.rindex2, self.rindex3], f["index"])
            if "middle" in f:
                applyFingerPose([self.rmiddle1, self.rmiddle2, self.rmiddle3], f["middle"])
            if "ring" in f:
                applyFingerPose([self.rring1, self.rring2, self.rring3], f["ring"])
            if "pinky" in f:
                applyFingerPose([self.rpinky1, self.rpinky2, self.rpinky3], f["pinky"])

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
            self.status_text.setText("No valid signs found in text")
            self.signing_complete = True
            return

        self.status_text.setText(f"Signing: {self.current_text}")
        self.pose_index = 0
        self.is_animating = True
        self.signing_complete = False

        # Pause media playback if needed during signing
        if self.media_control_active and self.media_state == "playing":
            self.pause_media()
        taskMgr.add(self.animateNextPose, "SignAnimation")

    def stopAnimation(self):
        """Stop the current sign animation task if running."""
        if self.is_animating:
            taskMgr.remove("SignAnimation")
            self.is_animating = False

    def slideArms(self):
        """
        Provide a simple sliding motion animation for fingerspelling transitions.
        """
        slide_distance = 0.5
        time = 0.2
        sequence = Sequence(
            LerpPosInterval(self.larm, time, self.larm.getPos()),
            LerpPosInterval(self.rarm, time, self.rarm.getPos() + LVecBase3f(-slide_distance, 0, 0)),
            LerpPosInterval(self.larm, time, self.larm.getPos()),
            LerpPosInterval(self.rarm, time, self.rarm.getPos())
        )
        sequence.start()

    def animateNextPose(self, task):
        """
        Animates each sign or letter in the expanded sequence, applying relevant poses.
        """
        if self.pose_index >= len(self.expanded_sequence):
            self.loadSignPoses("default")
            self.pose_index = 0
            self.is_animating = False
            self.status_text.setText("Animation Complete")
            self.current_pose = ""
            self.signing_complete = True
            # Resume media playback after signing if needed
            if self.media_control_active and self.media_state == "paused":
                self.resume_media()
            return Task.done

        pose_name = self.expanded_sequence[self.pose_index]
        # Avoid repeating the same pose for consecutive same letters
        if self.current_pose == pose_name and len(pose_name) == 1:
            self.slideArms()
            self.pose_index += 1
            return task.again

        self.current_pose = pose_name
        poses = self.gesture_data.get(pose_name)
        if not poses:
            self.pose_index += 1
            return task.again

        # Animation sequence for the current sign or letter
        sequence = []
        time = 0.05  # Duration per movement

        def addFingerLerp(hand_data, finger_map):
            if "fingers" not in hand_data:
                return
            for name, parts in finger_map.items():
                if name in hand_data["fingers"]:
                    for part, pose_data in zip(parts, hand_data["fingers"][name]):
                        sequence.append(LerpPosInterval(part, 0, LVecBase3f(*pose_data["pos"])))
                        sequence.append(LerpHprInterval(part, 0, LVecBase3f(*pose_data["hpr"])))

        def addHandAndFingers(pose):
            l = pose["leftHand"]
            r = pose["rightHand"]
            sequence.extend([
                LerpPosInterval(self.larm, time, LVecBase3f(*l["pos"])),
                LerpHprInterval(self.larm, time, LVecBase3f(*l["hpr"])),
                LerpPosInterval(self.rarm, time, LVecBase3f(*r["pos"])),
                LerpHprInterval(self.rarm, time, LVecBase3f(*r["hpr"]))
            ])
            addFingerLerp(l, {
                "thumb": [self.lthumb1, self.lthumb2],
                "index": [self.lindex1, self.lindex2, self.lindex3],
                "middle": [self.lmiddle1, self.lmiddle2, self.lmiddle3],
                "ring": [self.lring1, self.lring2, self.lring3],
                "pinky": [self.lpinky1, self.lpinky2, self.lpinky3]
            })
            addFingerLerp(r, {
                "thumb": [self.rthumb1, self.rthumb2],
                "index": [self.rindex1, self.rindex2, self.rindex3],
                "middle": [self.rmiddle1, self.rmiddle2, self.rmiddle3],
                "ring": [self.rring1, self.rring2, self.rring3],
                "pinky": [self.rpinky1, self.rpinky2, self.rpinky3]
            })

        if isinstance(poses, list):
            for pose in poses:
                addHandAndFingers(pose)
        else:
            addHandAndFingers(poses)

        Sequence(*sequence).start()
        self.status_text.setText(f"Signing: {self.current_text} ('{pose_name}')")
        task.delayTime = 1.5  # Wait before the next pose
        self.pose_index += 1
        return task.again

    def setup_media_control(self):
        """
        Set up background task for media control (play/pause).
        Takes no effect until activated.
        """
        taskMgr.add(self.media_control_task, "MediaControlTask")

    def toggle_media_control(self):
        """
        Enable or disable media control mode. Updates UI and resets timing.
        """
        try:
            self.media_control_active = not self.media_control_active

            if self.media_control_active:
                self.media_toggle_button["text"] = "Stop Media Control"
                self.media_toggle_button["frameColor"] = (0.9, 0.3, 0.3, 1)
                self.status_text.setText("Media control starting (switch to media tab)")
                self.last_media_action_time = time.time()
                self.media_state = "starting"
                print("Media control starting - switch to your media tab within 3 seconds!")
            else:
                self.media_toggle_button["text"] = "Start Media Control"
                self.media_toggle_button["frameColor"] = (0.3, 0.6, 0.9, 1)
                self.status_text.setText("Speech recognition active, media control inactive")
                self.media_state = "paused"
                print("Media control stopped")

        except Exception as e:
            print(f"Error toggling media control: {str(e)}")
            self.status_text.setText(f"Error: {str(e)}")

    def media_control_task(self, task):
        """
        Background task for alternating media play/pause based on timing.
        Only active if media_control_active flag is set.
        """
        if not self.media_control_active:
            return Task.cont
        if not self.signing_complete:
            return Task.cont  # Prevent media state change while signing

        current_time = time.time()
        elapsed = current_time - self.last_media_action_time

        if self.media_state == "starting" and elapsed >= 3:
            # After initial delay, start playback
            self.last_media_action_time = current_time
            self.media_state = "playing"
            self.status_text.setText("Media playing")

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
        self.status_text.setText("Media paused")

    def resume_media(self):
        """
        Simulate keyboard press to resume media and update internal state/UI.
        """
        self.simulate_space_press()
        self.last_media_action_time = time.time()
        self.media_state = "playing"
        self.status_text.setText("Media playing")

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
                print("Could not import win32com.client - media control may not work properly")
        else:
            try:
                import pyautogui
                pyautogui.press('space')
            except ImportError:
                print("Could not import pyautogui - media control may not work properly")

    def start_speech_recognition(self):
        """
        Launch the speech recognition processor and update UI state accordingly.
        """
        try:
            if not self.speech_processor:
                self.speech_processor = ContinuousSpeechGloss(callback=self.handle_speech_result)
            success = self.speech_processor.start()
            if success:
                self.speech_recognition_active = True
                self.speech_toggle_button["text"] = "Stop Speech Recognition"
                self.speech_toggle_button["frameColor"] = (0.9, 0.3, 0.3, 1)
                self.status_text.setText("Speech recognition active")
            else:
                self.status_text.setText("Error: Speech recognition failed to start")
        except Exception as e:
            error_msg = f"Error starting speech recognition: {str(e)}"
            print(error_msg)
            self.status_text.setText("Error: Failed to start speech recognition")

    def toggle_speech_recognition(self):
        """
        Toggle speech recognition on/off and update UI button and state.
        """
        if not self.speech_recognition_active:
            self.start_speech_recognition()
        else:
            try:
                if self.speech_processor:
                    success = self.speech_processor.stop()
                    if success:
                        self.speech_recognition_active = False
                        self.speech_toggle_button["text"] = "Start Speech Recognition"
                        self.speech_toggle_button["frameColor"] = (0.3, 0.6, 0.9, 1)
                        self.status_text.setText("Speech recognition inactive")
                    else:
                        self.status_text.setText("Error: Failed to stop speech recognition")
            except Exception as e:
                error_msg = f"Error stopping speech recognition: {str(e)}"
                print(error_msg)
                self.status_text.setText("Error: Failed to stop speech recognition")

    def handle_speech_result(self, text, gloss):
        """
        Called whenever new speech is recognized.
        Updates UI and triggers sign animation (unless an animation is already running).
        """
        if text and gloss:
            if self.is_animating:
                return  # Wait for current animation to complete
            self.start_animation(text)


# Main entry point, creates and runs the application
def run_app():
    app = SignLanguageApp()
    app.run()


if __name__ == "__main__":
    run_app()
