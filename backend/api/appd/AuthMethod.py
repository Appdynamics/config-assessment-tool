import asyncio
import ipaddress
import json
import logging
import re

import aiohttp
from uplink import AiohttpClient
from uplink.auth import BasicAuth, BearerToken

from api.Result import Result
from api.appd.AppDController import AppdController
from util.asyncio_utils import AsyncioUtils
from util.logging_utils import initLogging


class AuthMethod():

    controller: AppdController

    def __init__(self,
                 auth_method: str,
                 host,
                 port,
                 ssl=True,
                 account=None,
                 username=None,
                 password=None,
                 useProxy=True,
                 verifySsl=True,
                 controller: AppdController = None):

        self.auth_method = auth_method
        self.host = host
        self.port = port
        self.ssl = ssl
        self.account = account
        self.username = username
        self.password = password
        self.useProxy = useProxy
        self.verifySSL = verifySsl
        self.session = None
        connection_url = (f'{"https" if ssl else "http"}://{host}:{port}')


        # poor man's DI
        #TODO: replace with proper DI
        if controller is None:
            cookie_jar = aiohttp.CookieJar()
            try:
                if ipaddress.ip_address(host):
                    logging.warning(
                        f"Configured host {host} is an IP address. Consider using the DNS instead.")
                    logging.warning(
                        f"RFC 2109 explicitly forbids cookie accepting from URLs with IP address instead of DNS name.")
                    logging.warning(f"Using unsafe Cookie Jar.")
                    cookie_jar = aiohttp.CookieJar(unsafe=True)
            except ValueError:
                pass

            connector = aiohttp.TCPConnector(
                limit=AsyncioUtils.concurrentConnections, verify_ssl=True)

            self.session = aiohttp.ClientSession(connector=connector,
                                                 trust_env=True,
                                                 cookie_jar=cookie_jar)

            self.controller = AppdController(
                base_url=connection_url,
                client=AiohttpClient(session=self.session),
                session=self.session
            )
        else:
            self.session = controller.get_client_session()
            self.controller = controller

    async def loginBasicAuthentication(self):
        if self.auth_method.lower() == "basic":
            self.controller.session.auth = (
                BasicAuth(f"{self.username}@{self.account}",
                          self.password))

            try:
                response = await self.controller.login()
            except Exception as e:
                logging.error(
                    f"Controller login failed with {e}. Check username and password.")
                raise e

            if response.status_code != 200:
                err_msg = f"{self.host} - Controller login failed with " \
                             f"{response.status_code}. Check username and " \
                             f"password"
                logging.error(err_msg)
                return Result(
                    response,
                    Result.Error(err_msg),
                )
            try:
                jsessionid = \
                    re.search("JSESSIONID=(\\w|\\d)*", str(response.headers)).group(
                        0).split("JSESSIONID=")[1]
                self.controller.jsessionid = jsessionid
            except AttributeError:
                logging.info(
                    f"{self.host} - Unable to find JSESSIONID in login "
                    f"response. Please verify credentials.")
            try:
                xcsrftoken = re.search("X-CSRF-TOKEN=(\\w|\\d)*",
                                       str(response.headers)).group(0).split(
                    "X-CSRF-TOKEN=")[1]
                self.controller.xcsrftoken = xcsrftoken
            except AttributeError:
                logging.info(
                    f"{self.host} - Unable to find X-CSRF-TOKEN in login "
                    f"response. Please verify credentials.")

            if (self.controller.jsessionid is None or self.controller.xcsrftoken is
                    None):
                return Result(
                    response,
                    Result.Error(
                        f"{self.host} - Valid authentication headers not "
                        f"cached from previous login call. Please verify credentials."),
                )

            self.controller.session.headers["X-CSRF-TOKEN"] = (
                self.controller.xcsrftoken)
            self.controller.session.headers["Set-Cookie"] = (f"JSESSIONID="
                                                             f"{self.controller.jsessionid};X-CSRF-TOKEN={self.controller.xcsrftoken};")
            self.controller.session.headers["Content-Type"] = \
                "application/json;charset=UTF-8"

            return await self._validateBasicAuthPermissions(self.username)

    async def loginClientSecretOauthAuthentication(self):
        payload = {
            "grant_type": "client_credentials",
            "client_id": f"{self.username}@{self.account}",
            "client_secret": self.password,
        }

        try:
            response = await self.controller.loginOAuth(data=payload)
        except Exception as e:
            logging.error(f"Controller login failed with {e}. Check username and password.")
            return Result(None, Result.Error(f"{self.host} - Controller login failed with "))

        if response.status_code != 200:
            err_msg = f"{self.host} - Controller login failed with " \
                         f"{response.status_code}. Check username and " \
                         f"password 1."
            logging.error(err_msg)
            return Result(
                response,
                Result.Error(err_msg),
            )

        token_data = await response.json()
        token = token_data["access_token"]
        expires_in = token_data.get("expires_in", 300)
        self.controller.session.headers["Authorization"] = (f"Bearer"
                                                            f" {token}")
        self.controller.session.auth = BearerToken(token)

        return await self._validateAPIClientPermissions(self.username)

    async def loginTokenOauthAuthentication(self):
        self.controller.session.auth = BearerToken(self.password)
        return await self._validateAPIClientPermissions(self.username)

    async def authenticate(self):

        if self.auth_method.lower() == "basic":
            logging.info(f"Authenticating user {self.username}, using: "
                         f"<<<{self.auth_method} >>> " f"authentication for {self.host}")
            return await self.loginBasicAuthentication()


        if self.auth_method.lower() == "secret":
            logging.info(
                f"Authenticating user {self.username}, using API Client "
                f"authentication <<< {self.auth_method} >>> (Client "
                f"name/secret) for {self.host}")
            return await self.loginClientSecretOauthAuthentication()

        if self.auth_method.lower() == "token":
            logging.info(f"Authenticating user {self.username}, using API "
                         f"Client authentication <<< {self.auth_method} >>> "
                         f"(Temporary Access Token) for {self.host}")
            return await self.loginTokenOauthAuthentication()



    async def getAdminRoleId(self):
        response = await self.controller.getRoles()
        roles = await self.getResultFromResponse(response,
                                                 "Get roles")
        for role in roles.data:
            if role["name"] == "Account Administrator":
                return role["id"]

    async def _validateBasicAuthPermissions(self, username) -> Result:
        if username is None:
            msg = " - Username is required for basic auth authentication"
            logging.error(f"{self.host} {msg}")
            return Result(
                None,
                Result.Error(
                    f"{self.host} {msg}"),
            )

        response = await self.controller.getUsers()
        users = await (self.getResultFromResponse(response, "Get users"))

        errorMsg = f"{self.host} - Unable to validate permissions. Does {username} have Account Ownership role? "

        if users.error is not None:
            logging.error(errorMsg)
            return Result(
                None,
                Result.Error(errorMsg),
            )

        userID = next(user["id"] for user in users.data if
                      user["name"].lower() == username.lower())
        response = await self.controller.getUser(userID)
        result = await self.getResultFromResponse(response, "Get user")

        if result.error is not None:
            return Result(response.Error(errorMsg))

        adminRole = next((role for role in result.data["roles"] if role[
            "name"] == "super-admin"), None)
        if not adminRole:
            logging.error(errorMsg)
            return Result(None, Result.Error(errorMsg))
        # If adminRole is found, continue with this logic

        logging.debug(
            f"User permissions: Role: {adminRole['name']}, idx: {adminRole['id']}")
        logging.info(f"{self.host} - {username} admin "
                     f"role validated. User has Account Ownership role!")

        return Result(result.data, None)


    async def _validateAPIClientPermissions(self, username: str) -> Result:
        if username is None:
            error_msg = f"{self.host} - API Client username is required for token authentication."
            logging.error(error_msg)
            return Result(None, Result.Error(error_msg))

        raw_response = await self.controller.getApiClients()
        api_clients_result = await self.getResultFromResponse(raw_response, "Get API clients")

        if api_clients_result.error is not None:
            error_msg = (f"{self.host} - Failed to retrieve API clients: "
                         f"{api_clients_result.error.msg if api_clients_result.error.msg else 'Unknown error'}")
            logging.error(error_msg)
            return Result(raw_response, api_clients_result.error)

        if not api_clients_result.data: # Check if the data list is empty or None
            error_msg = (f"{self.host} - No API clients found. Unable to validate permissions for '{username}'. "
                         f"Please ensure API clients exist and {username} has the Account Ownership role.")
            logging.error(error_msg)
            return Result(raw_response, Result.Error(error_msg))

        found_api_client = None
        for client in api_clients_result.data:
            # Use .get() for safer access to dictionary keys, preventing KeyError
            if client.get("name") == username:
                found_api_client = client
                break # Found the client, no need to continue searching

        if found_api_client is None:
            error_msg = (f"{self.host} - Unable to validate user administrative roles/permission. "
                         f"API client '{username}' not found in the retrieved list.")
            logging.error(error_msg)
            return Result(raw_response, Result.Error(error_msg))

        account_role_ids = found_api_client.get("accountRoleIds", [])
        if not isinstance(account_role_ids, list):
            # Log a warning if the role IDs are not in the expected list format
            logging.warning(f"{self.host} - API client '{username}' has invalid 'accountRoleIds' format. Expected a list.")
            account_role_ids = [] # Treat as empty list if not valid

        admin_role_id = await self.getAdminRoleId()

        if admin_role_id is not None and admin_role_id in account_role_ids:
            logging.info(f"{self.host} - API client {username} admin role validated. User has Account Ownership role!")
            validated_roles = {"roles": [{"name": "Account Administrator"}]}
            return Result(validated_roles, None)
        else:
            error_msg = (f"{self.host} - Unable to validate user administrative roles/permission for '{username}'. "
                         f"Admin role ID could not be determined or the user lacks the required administrative role.")
            logging.error(error_msg)
            return Result(raw_response, Result.Error(error_msg))





    async def cleanup(self):
        logging.info(f"Cleaning up closing connection for this service using "
                     f"{self.host}")
        await self.session.close()

    async def getResultFromResponse(self, response, debugString,
                                    isResponseJSON=True,
                                    isResponseList=True) -> Result:
        body = (await response.content.read()).decode("ISO-8859-1")

        if response.status_code >= 400:
            msg = (f"{self.host} - {debugString} failed with code"
                   f":{response.status_code} body:{body}")
            try:
                responseJSON = json.loads(body)
                if "message" in responseJSON:
                    msg = (f"{self.host} - {debugString} failed with code"
                           f":{response.status_code} body:{responseJSON['message']}")
            except json.JSONDecodeError:
                pass
            logging.debug(msg)
            return Result([] if isResponseList else {},
                          Result.Error(f"{response.status_code}"))
        if isResponseJSON:
            try:
                return Result(json.loads(body), None)
            except json.JSONDecodeError:
                msg = (f"{self.host} - {debugString} failed to parse json from "
                       f"body. Returned code:{response.status_code} body:{body}")
                logging.error(msg)
                return Result([] if isResponseList else {}, Result.Error(msg))
        else:
            return Result(body, None)



async def main():

    authMethod = AuthMethod(
        auth_method="",
        host="",
        port=443,
        ssl=True,
        account="",
        username="",
        password="",
        useProxy=False,
        verifySsl=True
    )

    await authMethod.authenticate()
    await authMethod.cleanup()

    # inject controller
    host = ""
    ssl = True
    port = 443
    connection_url = (f'{"https" if ssl else "http"}://{host}:{port}')

    cookie_jar = aiohttp.CookieJar()
    try:
        if ipaddress.ip_address(host):
            logging.warning(
                f"Configured host {host} is an IP address. Consider using the DNS instead.")
            logging.warning(
                f"RFC 2109 explicitly forbids cookie accepting from URLs with IP address instead of DNS name.")
            logging.warning(f"Using unsafe Cookie Jar.")
            cookie_jar = aiohttp.CookieJar(unsafe=True)
    except ValueError:
        pass

    connector = aiohttp.TCPConnector(
        limit=AsyncioUtils.concurrentConnections, verify_ssl=True)

    client_session = aiohttp.ClientSession(connector=connector,
                                           trust_env=True,
                                           cookie_jar=cookie_jar)

    controller = AppdController(
        base_url=connection_url,
        client=AiohttpClient(session=client_session),
        session=client_session
    )

if __name__ == '__main__':
    initLogging(True)
    asyncio.run(main())
