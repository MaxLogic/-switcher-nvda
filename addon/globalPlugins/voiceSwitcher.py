import json
import os
import re
from copy import deepcopy

from logHandler import log

import addonHandler
import config
import gui
import scriptHandler
import synthDriverHandler
import ui
import wx
from globalPluginHandler import GlobalPlugin as BaseGlobalPlugin


addonHandler.initTranslation()


PERSISTENCE_FILE = os.path.join(config.getUserDefaultConfigPath(), "voice-switcher-presets.json")
APPLY_ORDER = ("voice", "language", "variant", "rate", "rateBoost", "pitch", "inflection", "volume")
FALLBACK_SETTING_IDS = APPLY_ORDER
NOISY_VOICE_TOKENS = (
	"desktop",
	"mobile",
	"multilingual",
	"online",
	"offline",
	"tts",
	"voice",
)
VOICE_VENDOR_PREFIXES = (
	"microsoft",
	"rhvoice",
	"vocalizer expressive",
	"vocalizer",
	"espeak ng",
	"espeak",
	"eloquence",
	"ibm",
	"acapela",
	"onecore",
	"sapi5",
	"sapi",
)
MISSING_SETTING = object()


def _get_synth():
	synth = synthDriverHandler.getSynth()
	if synth is None:
		return None
	return synth


def _voice_display_name(synth, voice_id, stored_voice_name=None):
	voices = getattr(synth, "availableVoices", {}) or {}
	voice_info = voices.get(voice_id)
	if voice_info is not None:
		return voice_info.displayName
	if stored_voice_name:
		return stored_voice_name
	if voice_id:
		return str(voice_id)
	return _("Unknown voice")


def _format_rate(rate):
	if rate is None:
		return _("unknown")
	return f"{rate}%"


def _clean_voice_name(voice_name):
	if not voice_name:
		return _("Unknown voice")

	name = voice_name.strip()
	name = re.sub(r"\(([^)]*(?:khz|hz|bit|sample)[^)]*)\)", "", name, flags=re.IGNORECASE)
	name = re.sub(r"\[[^\]]*(?:khz|hz|bit|sample)[^\]]*\]", "", name, flags=re.IGNORECASE)
	name = name.split(" - ", 1)[0].strip()
	name = re.sub(r"\b\d+(?:\.\d+)?\s*(?:khz|hz|bit)\b", "", name, flags=re.IGNORECASE)
	for prefix in VOICE_VENDOR_PREFIXES:
		if name.casefold().startswith(prefix + " "):
			name = name[len(prefix):].strip()
			break
	words = [word for word in name.split() if word.casefold() not in NOISY_VOICE_TOKENS]
	name = " ".join(words).strip(" -_,")
	return name or voice_name.strip()


def _serializable_value(value):
	return isinstance(value, (str, int, float, bool)) or value is None


def _get_setting_value(synth, setting_id, default=MISSING_SETTING):
	try:
		return getattr(synth, setting_id)
	except (AttributeError, NotImplementedError):
		return default


def _section_has_key(section, key):
	if hasattr(section, "isSet"):
		try:
			return section.isSet(key)
		except Exception:
			pass
	try:
		return key in section
	except Exception:
		return False


def _snapshot_section_value(section, key):
	if not _section_has_key(section, key):
		return None
	return deepcopy(section[key])


def _restore_profile_section_value(section, key, snapshot):
	if snapshot is None:
		if _section_has_key(section, key):
			try:
				del section[key]
			except Exception:
				pass
		return
	section[key].clear()
	section[key].update(deepcopy(snapshot))


def _restore_speech_section_value(section, key, snapshot):
	if snapshot is None:
		if _section_has_key(section, key):
			try:
				del section[key]
			except Exception:
				pass
		return
	section[key] = deepcopy(snapshot)
	if hasattr(section[key], "_cache"):
		section[key]._cache.clear()


def _capture_state(synth):
	state = {}
	for setting in getattr(synth, "supportedSettings", ()):
		setting_id = getattr(setting, "id", None)
		if not setting_id:
			continue
		value = _get_setting_value(synth, setting_id)
		if value is MISSING_SETTING:
			continue
		if _serializable_value(value):
			state[setting_id] = value
	for setting_id in FALLBACK_SETTING_IDS:
		if setting_id in state:
			continue
		value = _get_setting_value(synth, setting_id)
		if value is MISSING_SETTING:
			continue
		if _serializable_value(value):
			state[setting_id] = value
	return state


class PresetStore:
	def __init__(self, path):
		self.path = path
		self.presets = {}

	def load(self):
		try:
			with open(self.path, "r", encoding="utf-8") as stream:
				data = json.load(stream)
		except FileNotFoundError:
			return
		except Exception:
			log.warning("Failed to read voice preset store", exc_info=True)
			return

		presets = data.get("presets", {})
		if not isinstance(presets, dict):
			log.warning("Invalid voice preset store format")
			return

		valid_presets = {}
		for name, preset in presets.items():
			if not isinstance(name, str) or not isinstance(preset, dict):
				continue
			if not isinstance(preset.get("synth"), str):
				continue
			settings = preset.get("settings", {})
			if not isinstance(settings, dict):
				continue
			valid_presets[name] = {
				"name": name,
				"synth": preset["synth"],
				"settings": settings,
				"voice_name": preset.get("voice_name"),
				"rate": preset.get("rate"),
			}
		self.presets = valid_presets

	def save(self):
		os.makedirs(os.path.dirname(self.path), exist_ok=True)
		payload = {
			"version": 1,
			"presets": self.presets,
		}
		temp_path = f"{self.path}.tmp"
		with open(temp_path, "w", encoding="utf-8") as stream:
			json.dump(payload, stream, ensure_ascii=False, indent=2, sort_keys=True)
		os.replace(temp_path, self.path)

	def add(self, preset):
		self.presets[preset["name"]] = preset

	def delete(self, preset_name):
		if preset_name in self.presets:
			del self.presets[preset_name]

	def get_sorted_items(self):
		return sorted(self.presets.items(), key=lambda item: item[0].casefold())


def capture_current_preset(name):
	synth = _get_synth()
	if synth is None:
		raise RuntimeError(_("No active speech synthesizer."))

	settings = _capture_state(synth)
	voice_id = settings.get("voice", getattr(synth, "voice", None))
	rate = settings.get("rate", getattr(synth, "rate", None))
	voice_name = _voice_display_name(synth, voice_id)
	return {
		"name": name,
		"synth": synth.name,
		"settings": settings,
		"voice_name": voice_name,
		"rate": rate,
	}


def suggest_preset_name(preset):
	voice_name = preset.get("voice_name") or _("Unknown voice")
	return _clean_voice_name(voice_name)


def _apply_settings_to_config(synth_name, settings):
	profile_section = config.conf.profiles[0]["speech"][synth_name]
	profile_section.clear()
	profile_section.update(deepcopy(settings))
	config.conf["speech"][synth_name] = deepcopy(settings)
	config_section = config.conf["speech"][synth_name]
	if hasattr(config_section, "_cache"):
		config_section._cache.clear()


def apply_preset(preset):
	synth_name = preset.get("synth")
	settings = preset.get("settings", {})
	if not synth_name or not isinstance(settings, dict):
		ui.message(_("This preset is invalid."))
		return False

	log.info(f"Applying voice preset {preset.get('name', '<unnamed>')!r} for synth {synth_name!r}")
	log.info(f"Voice preset stored settings: {settings!r}")

	original_synth = _get_synth()
	original_name = original_synth.name if original_synth is not None else None
	speech_section = config.conf["speech"]
	profile_speech_section = config.conf.profiles[0]["speech"]
	original_target_config = None
	original_profile_target_config = None
	previous_target_config = None
	previous_profile_target_config = None

	try:
		if original_name:
			original_target_config = _snapshot_section_value(speech_section, original_name)
			original_profile_target_config = _snapshot_section_value(profile_speech_section, original_name)

		previous_target_config = _snapshot_section_value(speech_section, synth_name)
		previous_profile_target_config = _snapshot_section_value(profile_speech_section, synth_name)
		target_settings = deepcopy(settings)
		log.info(f"Voice preset target settings written to config: {target_settings!r}")
		_apply_settings_to_config(synth_name, target_settings)

		current_synth = _get_synth()
		if current_synth is None or current_synth.name != synth_name:
			if not synthDriverHandler.setSynth(synth_name):
				ui.message(_("Could not switch to synthesizer {synth}.").format(synth=synth_name))
				return False
			current_synth = _get_synth()
			if current_synth is None or current_synth.name != synth_name:
				ui.message(_("Could not switch to synthesizer {synth}.").format(synth=synth_name))
				return False
		else:
			current_synth.loadSettings(onlyChanged=True)
		if current_synth is None:
			raise RuntimeError("No active synth after applying preset")
		current_synth.saveSettings()
		log.info(
			f"Voice preset {preset.get('name', '<unnamed>')!r} applied successfully; "
			f"active voice={_get_setting_value(current_synth, 'voice', None)!r}"
		)
	except Exception:
		log.warning("Failed to apply voice preset", exc_info=True)
		_restore_speech_section_value(speech_section, synth_name, previous_target_config)
		_restore_profile_section_value(profile_speech_section, synth_name, previous_profile_target_config)
		if original_name is not None:
			_restore_speech_section_value(speech_section, original_name, original_target_config)
			_restore_profile_section_value(profile_speech_section, original_name, original_profile_target_config)
			rollback_synth = _get_synth()
			if rollback_synth is None or rollback_synth.name != original_name:
				synthDriverHandler.setSynth(original_name)
				rollback_synth = _get_synth()
			if rollback_synth is not None and rollback_synth.name == original_name:
				try:
					if _get_synth() is rollback_synth:
						rollback_synth.loadSettings(onlyChanged=True)
					rollback_synth.saveSettings()
				except Exception:
					log.warning("Failed to persist rolled-back synth state", exc_info=True)
		ui.message(_("The preset could not be applied."))
		return False
	return True


class PresetManagerDialog(wx.Dialog):
	def __init__(self, parent, store, initiallySelectedPresetName=None, onPresetLoaded=None):
		super().__init__(
			parent,
			title=_("Voice presets"),
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self.store = store
		self._entries = []
		self._initiallySelectedPresetName = initiallySelectedPresetName
		self._onPresetLoaded = onPresetLoaded

		mainSizer = wx.BoxSizer(wx.VERTICAL)
		sizerHelper = gui.guiHelper.BoxSizerHelper(self, orientation=wx.VERTICAL)
		self.listBox = sizerHelper.addLabeledControl(
			_("Saved &presets"),
			wx.ListBox,
			style=wx.LB_SINGLE | wx.LB_HSCROLL,
			size=(560, 240),
		)
		self.listBox.Bind(wx.EVT_LISTBOX_DCLICK, self.onLoad)
		self.listBox.Bind(wx.EVT_LISTBOX, self.onSelectionChanged)

		buttonHelper = gui.guiHelper.ButtonHelper(orientation=wx.HORIZONTAL)
		self.loadButton = buttonHelper.addButton(self, wx.NewIdRef(), _("&Load"))
		self.saveButton = buttonHelper.addButton(self, wx.NewIdRef(), _("Save current &setting"))
		self.renameButton = buttonHelper.addButton(self, wx.NewIdRef(), _("&Rename"))
		self.deleteButton = buttonHelper.addButton(self, wx.NewIdRef(), _("&Delete"))
		self.closeButton = buttonHelper.addButton(self, wx.ID_CANCEL, _("&Close"))

		sizerHelper.addItem(buttonHelper)

		self.loadButton.Bind(wx.EVT_BUTTON, self.onLoad)
		self.saveButton.Bind(wx.EVT_BUTTON, self.onSaveCurrent)
		self.renameButton.Bind(wx.EVT_BUTTON, self.onRename)
		self.deleteButton.Bind(wx.EVT_BUTTON, self.onDelete)
		self.closeButton.Bind(wx.EVT_BUTTON, self.onClose)
		self.Bind(wx.EVT_CHAR_HOOK, self.onDialogCharHook)

		self.EscapeId = wx.ID_CANCEL
		self.SetAffirmativeId(self.loadButton.Id)
		self.loadButton.SetDefault()
		mainSizer.Add(sizerHelper.sizer, border=gui.guiHelper.BORDER_FOR_DIALOGS, flag=wx.ALL)
		self.SetSizer(mainSizer)
		mainSizer.Fit(self)
		self.refresh_list(select_name=self._initiallySelectedPresetName)
		self.CenterOnScreen()
		self.Bind(wx.EVT_SHOW, self.onShow)

	def refresh_list(self, select_name=None):
		self.listBox.Clear()
		self._entries = []
		for name, preset in self.store.get_sorted_items():
			self._entries.append((name, preset))
			self.listBox.Append(self.format_entry(name, preset))

		if not self._entries:
			self._sync_buttons()
			return

		selected_index = 0
		if select_name is not None:
			for index, (name, _) in enumerate(self._entries):
				if name == select_name:
					selected_index = index
					break
		self.listBox.SetSelection(selected_index)
		self._sync_buttons()

	def onShow(self, event):
		if event.IsShown():
			wx.CallAfter(self.Raise)
			wx.CallAfter(self.focusPresetList)
		event.Skip()

	def focusPresetList(self):
		if self._entries:
			index = self.listBox.GetSelection()
			if index == wx.NOT_FOUND:
				index = 0
				self.listBox.SetSelection(index)
			self.listBox.SetFocus()
		else:
			self.listBox.SetFocus()

	def format_entry(self, name, preset):
		voice_name = preset.get("voice_name") or _("Unknown voice")
		rate = _format_rate(preset.get("rate"))
		return _("{name} | Voice: {voice} | Speed: {speed}").format(
			name=name,
			voice=voice_name,
			speed=rate,
		)

	def _selected_entry(self):
		index = self.listBox.GetSelection()
		if index == wx.NOT_FOUND or index < 0 or index >= len(self._entries):
			return None
		return self._entries[index]

	def _sync_buttons(self):
		has_selection = self._selected_entry() is not None
		self.loadButton.Enable(has_selection)
		self.renameButton.Enable(has_selection)
		self.deleteButton.Enable(has_selection)

	def onSelectionChanged(self, event):
		self._sync_buttons()
		event.Skip()

	def onDialogCharHook(self, event):
		key = event.GetKeyCode()
		if key in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
			self.onLoad(None, closeAfterLoad=True)
			return
		if key == wx.WXK_F2:
			self.onRename(None)
			return
		if key == wx.WXK_DELETE:
			self.onDelete(None)
			return
		event.Skip()

	def onLoad(self, event, closeAfterLoad=False):
		selected = self._selected_entry()
		if selected is None:
			log.info("Load requested without a selected preset")
			ui.message(_("Select a preset first."))
			return
		name, preset = selected
		log.info(f"Load requested for preset {name!r}")
		if apply_preset(preset):
			ui.message(_("Loaded preset {name}.").format(name=name))
			if self._onPresetLoaded is not None:
				self._onPresetLoaded(name)
			if closeAfterLoad:
				if self.IsModal():
					self.EndModal(wx.ID_OK)
				else:
					self.Close()

	def onSaveCurrent(self, event):
		try:
			preset = capture_current_preset("")
		except RuntimeError as error:
			ui.message(str(error))
			return

		default_name = suggest_preset_name(preset)
		dialog = wx.TextEntryDialog(
			self,
			_("Enter a name for this preset."),
			_("Save current setting"),
			value=default_name,
		)
		try:
			if dialog.ShowModal() != wx.ID_OK:
				return
			name = dialog.GetValue().strip()
		finally:
			dialog.Destroy()

		if not name:
			return

		replacing_existing = name in self.store.presets
		previous_preset = deepcopy(self.store.presets.get(name)) if name in self.store.presets else None
		preset["name"] = name
		self.store.add(preset)
		if not self._save_store():
			if previous_preset is None:
				self.store.delete(name)
			else:
				self.store.add(previous_preset)
			return
		self.refresh_list(select_name=name)
		ui.message(_("Preset updated.") if replacing_existing else _("Preset saved."))

	def onRename(self, event):
		selected = self._selected_entry()
		if selected is None:
			return
		old_name, preset = selected
		dialog = wx.TextEntryDialog(
			self,
			_("Enter a new name for this preset."),
			_("Rename preset"),
			value=old_name,
		)
		try:
			if dialog.ShowModal() != wx.ID_OK:
				return
			new_name = dialog.GetValue().strip()
		finally:
			dialog.Destroy()

		if not new_name or new_name == old_name:
			return

		replaced_preset = deepcopy(self.store.presets.get(new_name)) if new_name in self.store.presets else None
		if replaced_preset is not None and gui.messageBox(
			_("A preset named {name} already exists. Replace it?").format(name=new_name),
			_("Replace preset"),
			wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION,
			self,
		) != wx.YES:
			return

		original_preset = deepcopy(preset)
		renamed_preset = deepcopy(preset)
		renamed_preset["name"] = new_name
		self.store.delete(old_name)
		self.store.add(renamed_preset)
		if not self._save_store():
			self.store.delete(new_name)
			self.store.add(original_preset)
			if replaced_preset is not None:
				self.store.add(replaced_preset)
			return
		self.refresh_list(select_name=new_name)
		ui.message(_("Preset renamed."))

	def onDelete(self, event):
		selected = self._selected_entry()
		if selected is None:
			return
		name, _preset = selected
		if (
			gui.messageBox(
				_("Delete preset {name}? This cannot be undone.").format(name=name),
				_("Delete preset"),
				wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION,
				self,
			)
			!= wx.YES
		):
			return
		previous_preset = deepcopy(self.store.presets.get(name))
		self.store.delete(name)
		if not self._save_store():
			if previous_preset is not None:
				self.store.add(previous_preset)
			return
		self.refresh_list()
		ui.message(_("Preset deleted."))

	def onClose(self, event):
		if self.IsModal():
			self.EndModal(wx.ID_CANCEL)
		else:
			self.Close()

	def _save_store(self):
		try:
			self.store.save()
		except Exception:
			gui.messageBox(
				_("Could not save presets to disk. Check your NVDA configuration folder permissions."),
				_("Save error"),
				wx.OK | wx.ICON_ERROR,
				self,
			)
			return False
		return True


class GlobalPlugin(BaseGlobalPlugin):
	scriptCategory = _("Voice Switcher")

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		self.store = PresetStore(PERSISTENCE_FILE)
		self.store.load()
		self.lastLoadedPresetName = None

	@scriptHandler.script(
		description=_("Opens the voice preset manager."),
		gesture="kb:nvda+w",
	)
	def script_showPresetManager(self, gesture):
		preferred_name = self.lastLoadedPresetName if self.lastLoadedPresetName in self.store.presets else None
		dialog = PresetManagerDialog(
			gui.mainFrame,
			self.store,
			initiallySelectedPresetName=preferred_name,
			onPresetLoaded=self._rememberLoadedPreset,
		)
		gui.runScriptModalDialog(dialog)

	def _rememberLoadedPreset(self, presetName):
		self.lastLoadedPresetName = presetName
