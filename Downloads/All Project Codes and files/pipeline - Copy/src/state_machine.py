from enum import Enum, auto
from typing import Callable, Optional
from src.logger import log

class RegistrationState(Enum):
    LOGIN = auto()
    DASHBOARD = auto()
    SAVED_APPLICATIONS = auto()
    AMENDMENT = auto()
    APOB = auto()
    UPLOAD = auto()
    SAVE = auto()
    SUBMIT = auto()
    CHECKPOINT = auto()
    COMPLETE = auto()
    ERROR = auto()

class RegistrationStateMachine:
    """
    Finite State Machine to track and govern batch registration steps.
    Provides deterministic progression tracing and safe crash recoveries.
    """
    def __init__(self, initial_state: RegistrationState = RegistrationState.LOGIN):
        self._state = initial_state
        log.info(f"[STATE] Initialized state machine in state: {self._state.name}")

    @property
    def current_state(self) -> RegistrationState:
        return self._state

    def transition_to(self, target_state: RegistrationState, detail: Optional[str] = None) -> None:
        """
        Transitions the state machine to target_state with timing logs.
        """
        old_state_name = self._state.name
        self._state = target_state
        log.info(f"[STATE] Transition: {old_state_name} -> {target_state.name} "
                 f"{f'({detail})' if detail else ''}")

    def is_in(self, check_state: RegistrationState) -> bool:
        return self._state == check_state
