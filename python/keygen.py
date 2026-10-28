"""Generate keys elsewhere; Excel never needs Python. No registry or short-key lookup."""
import argparse
import base64
import json
import secrets
from pathlib import Path
from cryptography.hazmat.primitives.asymmetric.x25519 import X25519PrivateKey

DEMO_P = 154742504910672534362390399
DEMO_Q = 77371252455336267181195199
DEMO_G = 4

def encode(data):
    return base64.urlsafe_b64encode(data).decode('ascii').rstrip('=')

def generate(mode):
    if mode == 'demo':
        private = secrets.randbelow(DEMO_Q-2)+2
        public = pow(DEMO_G, private, DEMO_P)
        return encode(public.to_bytes(11,'little')), encode(private.to_bytes(11,'little'))
    if mode == 'production':
        private = X25519PrivateKey.generate()
        return encode(private.public_key().public_bytes_raw()), encode(private.private_bytes_raw())
    raise ValueError('Unknown mode')

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--mode',required=True,choices=['demo','production'])
    p.add_argument('--output',required=True,type=Path,help='New private key JSON file; do not share')
    a=p.parse_args()
    public,private=generate(a.mode)
    with a.output.open('x',encoding='utf-8') as f:
        json.dump({'mode':a.mode,'public_key':public,'private_key':private},f,indent=2)
        f.write('\n')
    print('Mode:',a.mode)
    print('Public key:',public)
    print('Private key saved to:',a.output)
    if a.mode=='demo':
        print('DEMONSTRATION ONLY: tiny finite-field DH group; not for confidential data.')

if __name__=='__main__': main()
