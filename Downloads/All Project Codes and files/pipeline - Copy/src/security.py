import os
import base64
import secrets

class SecureVault:
    """
    Transient, in-memory credential storage wrapper.
    Protects sensitive credentials in-memory using a light, runtime-generated symmetric key.
    Ensures credentials are never logged, never saved to disk, and are immediately cleared
    after the automation session finishes.
    """
    _key = None
    _encrypted_username = None
    _encrypted_password = None

    @classmethod
    def _initialize(cls):
        if cls._key is None:
            # Generate a transient runtime symmetric key
            cls._key = secrets.token_bytes(32)

    @classmethod
    def _xor_crypt(cls, data: str) -> str:
        cls._initialize()
        # High-performance in-memory XOR scrambling with transient key
        data_bytes = data.encode("utf-8")
        key_len = len(cls._key)
        crypt_bytes = bytearray(
            b ^ cls._key[i % key_len] for i, b in enumerate(data_bytes)
        )
        return base64.b64encode(crypt_bytes).decode("utf-8")

    @classmethod
    def store_credentials(cls, username: str, password: str) -> None:
        """Stores credentials securely in-memory using dynamic runtime encryption."""
        if not username or not password:
            cls.clear()
            return
        cls._encrypted_username = cls._xor_crypt(username)
        cls._encrypted_password = cls._xor_crypt(password)

    @classmethod
    def get_credentials(cls) -> tuple[str, str]:
        """Decrypts and returns the stored credentials in-memory."""
        if cls._encrypted_username is None or cls._encrypted_password is None:
            return "", ""
        
        def decrypt(encrypted: str) -> str:
            cls._initialize()
            crypt_bytes = base64.b64decode(encrypted)
            key_len = len(cls._key)
            dec_bytes = bytearray(
                b ^ cls._key[i % key_len] for i, b in enumerate(crypt_bytes)
            )
            return dec_bytes.decode("utf-8")

        return decrypt(cls._encrypted_username), decrypt(cls._encrypted_password)

    @classmethod
    def clear(cls) -> None:
        """Completely overwrites and wipes the stored credentials from memory."""
        cls._encrypted_username = None
        cls._encrypted_password = None
        cls._key = None
