"""Temporary test login; replace credential validation with the college database later."""
import hmac
import os
import secrets
import time
from threading import Lock
from fastapi import HTTPException, Request, Response
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field

COOKIE = 'attendance_session'
sessions = {}
attempts = {}
lock = Lock()

class Credentials(BaseModel):
    username: str = Field(max_length=200)
    password: str = Field(max_length=500)
    remember: bool = False

def install_auth(app):
    @app.middleware('http')
    async def protect_api(request: Request, call_next):
        path = request.url.path
        if path.startswith('/api/'):
            if request.method not in ('GET', 'HEAD', 'OPTIONS') and request.headers.get('x-attendance-request') != '1':
                return JSONResponse({'detail': 'Invalid request.'}, status_code=403)
            if path not in ('/api/auth/login', '/api/auth/logout', '/api/health'):
                with lock:
                    entry = sessions.get(request.cookies.get(COOKIE))
                if not entry or entry['expires'] <= time.time():
                    return JSONResponse({'detail': 'Please sign in to continue.'}, status_code=401)
                request.state.username = entry['username']
        return await call_next(request)

    @app.post('/api/auth/login')
    def login(credentials: Credentials, request: Request, response: Response):
        now = time.time()
        address = request.client.host if request.client else 'local'
        with lock:
            for key in list(attempts):
                if attempts[key][1] <= now:
                    del attempts[key]
            count, until = attempts.get(address, (0, now + 300))
            if count >= 10:
                raise HTTPException(429, 'Too many attempts. Try again in five minutes.')
            attempts[address] = (count + 1, until)
        username = os.getenv('ATTENDANCE_TEST_USERNAME', 'admin')
        password = os.getenv('ATTENDANCE_TEST_PASSWORD', 'iare@cdc')
        valid_user = hmac.compare_digest(credentials.username.encode(), username.encode())
        valid_password = hmac.compare_digest(credentials.password.encode(), password.encode())
        if not (valid_user and valid_password):
            raise HTTPException(401, 'Incorrect username or password.')
        lifetime = 7 * 86400 if credentials.remember else 8 * 3600
        token = secrets.token_urlsafe(32)
        with lock:
            attempts.pop(address, None)
            for key in list(sessions):
                if sessions[key]['expires'] <= now:
                    del sessions[key]
            sessions.pop(request.cookies.get(COOKIE), None)
            sessions[token] = {'username': username, 'expires': now + lifetime}
        response.set_cookie(COOKIE, token, httponly=True, samesite='strict', secure=os.getenv('COOKIE_SECURE') == '1', max_age=lifetime if credentials.remember else None)
        response.headers['Cache-Control'] = 'no-store'
        return {'username': username}

    @app.get('/api/auth/session')
    def session(request: Request, response: Response):
        response.headers['Cache-Control'] = 'no-store'
        return {'username': request.state.username}

    @app.post('/api/auth/logout')
    def logout(request: Request, response: Response):
        with lock:
            sessions.pop(request.cookies.get(COOKIE), None)
        response.delete_cookie(COOKIE)
        return {'ok': True}
