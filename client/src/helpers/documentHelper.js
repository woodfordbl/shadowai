/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

export function filterUserProfileInfo(result) {
  let userProfileInfo = [];
  userProfileInfo.push(result["mail"]);
  return userProfileInfo;
}
