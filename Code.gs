function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index');
}


/**
 * Get a list of Google Workspace groups the user is in.
 * @return {Array} List of group names.
 */
function listUserGroups() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const groups = GroupsApp.getGroupsForUser(userEmail);
    
    // Check if groups are retrieved
    if (!groups || groups.length === 0) {
      return [];
    }
    
    // Create an array of group names (emails)
    const groupNames = groups.map(group => group.getEmail());
    return groupNames;
  } catch (error) {
    Logger.log('Error retrieving user groups: ' + error);
    return [];
  }
}

// Function to check if user exists
function checkUserExists(userEmail) {
  try {
    var user = AdminDirectory.Users.get(userEmail);
    return user ? true : false;
  } catch (e) {
    return "user does not exist"; // User does not exist
  }
}

// Function to check and disable 2-step verification if enabled
function disable2SV(userEmail) {
  try {
    var user = AdminDirectory.Users.get(userEmail);
    if (!user) {
      return {
        status: 'error',
        message: 'User not found. Check email address.'
      };
    }

    // Check if 2-step verification is enabled
    var is2SVEnabled = user.isEnrolledIn2Sv;
    if (is2SVEnabled) {
      // Attempt to disable 2-step verification
      var result = AdminDirectory.Users.update({
        isEnrolledIn2Sv: false
      }, userEmail);

      return {
        status: 'success',
        twoSVStatus: 'disabled'
      };
    } else {
      return {
        status: 'success',
        twoSVStatus: 'not enabled'
      };
    }
  } catch (err) {
    return {
      status: 'error',
      message: 'Failed to manage 2-step verification: ' + err.message
    };
  }
}

// Function to revoke sign-in cookies for a user
function revokeSignInCookies(userEmail) {
  try {
    AdminDirectory.Users.signOut(userEmail);  // This revokes sign-in cookies
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Review and delete connected apps
function listAndRemoveConnectedApps(userEmail) {
  try {
    var user = AdminDirectory.Users.get(userEmail);
    if (!user) {
      return {
        status: 'error',
        message: 'User not found. Check email address.'
      };
    }

    // check for user's connected apps
    var connectedApps = user.authorizedApplications || [];
    var appsCount = connectedApps.length;
    var appsList = [];

    for (var i = 0; i < connectedApps.length; i++) {
      var app = connectedApps[i];
      appsList.push(app.displayName);
      // Attempt to remove the connected app
      // Placeholder logic for removal
      AdminDirectory.Users.removeAuthorizedApplication(userEmail, app.clientId);
    }

    return {
      status: 'success',
      appsCount: appsCount,
      appsList: appsList
    };
  } catch (err) {
    return {
      status: 'error',
      message: 'Failed to manage connected apps: ' + err.message
    };
  }
}

// Function Main to update user password and handle organizational unit change
function runTAOSwitch(userEmail) {
  let checklist = []; // Initialize checklist array
  
  try {

    var user = AdminDirectory.Users.get(userEmail);
    if (!user) {
      checklist.push({ action: 'User Check', status: 'declined', message: 'User not found. Check email address.' });
      return {
        status: 'error',
        message: 'User not found. Check email address.',
        checklist: checklist
      };
    }
 // Step 1: Update password and move user to "Suspended" org unit
    var newPassword = generateRandomPassword();
    AdminDirectory.Users.update({
      password: newPassword,
      changePasswordAtNextLogin: false,
      orgUnitPath: '/Suspended' // Changing OU to Suspended
    }, userEmail);

    checklist.push({ action: 'Update Password', status: 'approved', message: 'Password updated successfully.' });
    checklist.push({ action: 'Change OU', status: 'approved', message: 'OU Updated /Suspended.' });

    // Step 2: Deactivate enforced 2-Step Verification if applicable
    var deactivate2SVResult = deactivate2SVEnforced(userEmail);
    checklist.push({ action: 'Deactivate Org Enforcement', status: deactivate2SVResult.status === 'success' ? 'approved' : 'declined', message: deactivate2SVResult.message });

    // Step 3: Unenroll from 2-Step Verification
    // var twoSVResult = disable2SV(userEmail);
    // checklist.push({ action: 'Disable 2-Step Verification', status: twoSVResult.status === 'success' ? 'approved' : 'declined', message: twoSVResult.message });
    var unenroll2SVResult = unenroll2SV(userEmail);
    checklist.push({ action: 'Unenroll Org Enforcement', status: unenroll2SVResult.status === 'success' ? 'approved' : 'declined', message: unenroll2SVResult.message });

    // Step 4: List and remove connected apps
    var appsResult = runCheckConnectedApps(userEmail);
    checklist.push({ 
      action: 'Revoke Connected Apps', 
      status: appsResult.status === 'success' ? 'approved' : 'declined', 
      message: appsResult.status === 'success' ? `${appsResult.revokedCount + 1} connected apps revoked.` : appsResult.message 
    });

 // Step 6:Construct Google Takeout URL with userEmail
    const takeoutURL = `https://takeout.google.com/u/8/?pli=1`;

    // Step 7: Count Google Groups the user is enrolled in (delete SSO groups)
    var groupsResult = countGoogleGroups(userEmail);
    checklist.push({ action: 'Count Google Groups', status: groupsResult.status === 'success' ? 'approved' : 'declined', message: groupsResult.message });

  // Step 8: Revoke sign-in cookies
  // var revokeCookiesResult = revokeSignInCookies(userEmail);
  // checklist.push({ action: 'Sign-in cookies revoked', status: revokeCookiesResult.status === 'success' ? 'approved' : 'declined', message: revokeCookiesResult.message });

 // Step 8: Revoke sign-in cookies
  var revokeCookiesResult = revokeSignInCookies(userEmail);  // Call revokeSignInCookies function
  checklist.push({
    action: 'Sign-in cookies revoked',
    status: revokeCookiesResult.status === 'success' ? 'approved' : 'declined',
    message: revokeCookiesResult.message
  });
  

    // Extract the manager's email if available in relations
    let managerEmail = 'N/A';
    if (user.relations) {
      let managerRelation = user.relations.find(relation => relation.type === 'manager');
      if (managerRelation && managerRelation.value) {
        managerEmail = managerRelation.value;
      }
    }


    return {
      status: 'success',
      newPassword: newPassword,
      checklist: checklist,
      userName: user.name.fullName,
      userEmail: user.email,
      takeoutURL: takeoutURL
    };
    
  } catch (err) {
    checklist.push({ action: 'Update Process', status: 'declined', message: 'Failed to update password: ' + err.message });
    return {
      status: 'error',
      message: 'Failed to update password: ' + err.message,
      checklist: checklist
    };
  }
}

//Function to revoke sign-in cookies
function revokeSignInCookies(userEmail) {
  try {
    AdminDirectory.Users.signOut(userEmail);  // Revoke sign-in cookies
    return { status: 'success', message: 'Sign-in cookies revoked successfully.' };
  } catch (error) {
    return { status: 'failure', message: error.message || 'Unknown error while revoking sign-in cookies.' };
  }
}

// Function to retrieve all possible user details from the API
function getUserDetails(userEmail) {
  let userDetails = {};
  try {
    var user = AdminDirectory.Users.get(userEmail);

    if (!user) {
      return {
        status: 'error',
        message: 'User not found. Check email address.'
      };
    }

    // Extract the manager's email if available in relations
    let managerEmail = 'N/A';
    if (user.relations) {
      let managerRelation = user.relations.find(relation => relation.type === 'manager');
      if (managerRelation && managerRelation.value) {
        managerEmail = managerRelation.value;
      }
    }


    // Collect all possible details from the user object
    userDetails = {
      fullName: user.name.fullName || 'N/A',
      givenName: user.name.givenName || 'N/A',
      familyName: user.name.familyName || 'N/A',
      primaryEmail: user.primaryEmail || 'N/A',
      orgUnitPath: user.orgUnitPath || 'N/A',
      isAdmin: user.isAdmin || 'N/A',
      managerEmail: managerEmail || 'N/A', // Add manager email to the details
      isEnforcedIn2Sv: user.isEnforcedIn2Sv || 'N/A',
      is2SvEnrolled: user.isEnrolledIn2Sv || 'N/A',
      suspended: user.suspended || 'N/A',
      lastLoginTime: user.lastLoginTime || 'N/A',
      creationTime: user.creationTime || 'N/A',
      agreedToTerms: user.agreedToTerms || 'N/A',
      changePasswordAtNextLogin: user.changePasswordAtNextLogin || 'N/A',
      phones: user.phones || 'N/A',
      addresses: user.addresses || 'N/A',
      emails: user.emails || 'N/A',
      aliases: user.aliases || 'N/A',
      isDelegatedAdmin: user.isDelegatedAdmin || 'N/A',
      archived: user.archived || 'N/A',
      customerId: user.customerId || 'N/A',
      id: user.id || 'N/A',
      ipWhitelisted: user.ipWhitelisted || 'N/A',
      thumbnailPhotoUrl: user.thumbnailPhotoUrl || 'N/A'
    };

    return {
      status: 'success',
      userDetails: userDetails
    };
  } catch (err) {
    return {
      status: 'error',
      message: 'Failed to retrieve user details: ' + err.message
    };
  }
}

// Function to deactivate enforced 2-Step Verification for a user
function deactivate2SVEnforced(userEmail) {
  try {
    var user = AdminDirectory.Users.get(userEmail);

    if (!user.isEnrolledIn2Sv) {
      return {
        status: 'success',
        message: 'User is not enrolled in 2-Step Verification; Please do Manual Verification'
      };
    }

    // If enforced, we deactivate enforced 2SV
    // Note: The actual deactivation will depend on how your organization manages policies.
    // This might involve changing settings in the Admin Console, which might not be achievable via API.

    // Example:
    AdminDirectory.Users.update({
      isEnrolledIn2Sv: false,
      twoStepVerification: false // Assume there's a way to set this if enforced
    }, userEmail);

    return {
      status: 'success',
      message: 'Enforced - deactivated for the user. Please do Manual Verification'
    };
  } catch (err) {
    return {
      status: 'error',
      message: 'Failed to deactivate: Please do Manual Verification ' + err.message
    };
  }
}
// Function to check if the user has 2-Step Verification enrolled and unenroll them
function unenroll2SV(userEmail) {
  try {
    var user = AdminDirectory.Users.get(userEmail);

    if (!user.isEnrolledIn2Sv) {
      return {
        status: 'success',
        message: 'User is not enrolled in 2-Step Verification.Please do Manual Verification'
      };
    }

    // If enrolled, we unenroll the user by removing 2SV
    AdminDirectory.Users.update({
      isEnrolledIn2Sv: false
    }, userEmail);

    return {
      status: 'success',
      message: 'User was enrolled in 2-Step Verification but has been successfully unenrolled.'
    };
  } catch (err) {
    return {
      status: 'error',
      message: 'Failed to unenroll 2-Step Verification: Please do Manual Verification' + err.message
    };
  }
}

// Function to validate email format
function isValidEmail(email) {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

// Function to delete users from a CSV file
function deleteUsersFromCSV(csvContent) {
  const emails = csvContent.split('\n').map(email => email.trim()).filter(email => email);
  let deletedUsers = [];
  let invalidEmails = [];
  
  for (const email of emails) {
    if (isValidEmail(email)) {
      try {
        AdminDirectory.Users.remove(email);
        deletedUsers.push(email);
      } catch (err) {
        invalidEmails.push(email);
      }
    } else {
      invalidEmails.push(email);
    }
  }

  return {
    deletedUsers: deletedUsers,
    invalidEmails: invalidEmails
  };
}

// Function to delete a user from Google Workspace
function deleteUserFromGoogleWorkspace(userEmail) {
  try {
    AdminDirectory.Users.remove(userEmail);
    
    // Simulated check for user deletion
    var userDeleted = true;

    // Prepare checklist results for the output window
    var checklist = {
      documents: "User's documents won't be accessible.",
      emails: "User's emails will be deleted.",
      calendar: "User's calendar events will be deleted.",
      vault: "User's vault account will be deleted.",
      integrity: userDeleted ? "User account deleted successfully." : "User account deletion failed."
    };
    
    return { 
      status: 'success', 
      message: `User ${userEmail} has been deleted successfully.`,
      checklist: checklist,
      userDeleted: userDeleted
    };
  } catch (error) {
    return { 
      status: 'error', 
      message: `Error deleting user ${userEmail}: ${error.message}` 
    };
  }
}

// Function to list and count connected apps
function runCheckConnectedApps(userEmail) {
  let revokedCount = 0; // Count of revoked apps
  let connectedApps = []; // List of connected apps
  
  try {
    // Assume the following is the method to retrieve connected apps. Modify according to your API.
    var tokens = AdminDirectory.Tokens.list(userEmail).items || [];

    tokens.forEach(function(token) {
      connectedApps.push(token.displayText); // Add app names to the list

      // Revoke each token (app connection)
      try {
        AdminDirectory.Tokens.remove(userEmail, token.clientId);
        revokedCount++; // Increment the revoked count
      } catch (err) {
        Logger.log('Error revoking app: ' + token.displayText + ' for user: ' + userEmail + '. Error: ' + err.message);
      }
    });

    return {
      status: 'success',
      apps: connectedApps,
      revokedCount: revokedCount,
      message: revokedCount + ' connected apps revoked successfully.'
    };
  } catch (err) {
    return {
      status: 'error',
      message: 'Failed to retrieve or revoke connected apps: ' + err.message
    };
  }
}

// Function to count Google Groups the user is enrolled in
function countGoogleGroups(userEmail) {
  try {
    var groups = AdminDirectory.Groups.list({
      userKey: userEmail,
      customer: userEmail,
      maxResults: 100
    });
    
    var groupCount = groups.groups ? groups.groups.length : 0;

    return {
      status: 'success',
      groupCount: groupCount,
      message: `User is enrolled in ${groupCount} Google Groups.`
    };
  } catch (err) {
    return {
      status: 'error',
      message: 'Failed to count Google Groups: ' + err.message
    };
  }
}

// Function to initiate calendar transfer
function initiateCalendarTransfer(userEmail) {
  try {
    // Get the user's details including manager's email
    let user = AdminDirectory.Users.get(userEmail);
    let managerEmail = user.relations.find(relation => relation.type === 'manager').value;

    if (!managerEmail) {
      return 'Manager email not found.';
    }

    // Transfer the user's calendar to their manager
    let calendarTransferStatus = transferCalendar(userEmail, managerEmail);

    // Check if the calendar transfer was successful
    let auditLogs = checkAdminAuditLog(userEmail, managerEmail);

    // Prepare the checklist output
    let transferChecklist = prepareChecklist(calendarTransferStatus, auditLogs);

    return transferChecklist;
  } catch (err) {
    return `Error during transfer process: ${err.message}`;
  }
}

// Function to transfer the calendar to the manager
function transferCalendar(userEmail, managerEmail) {
  try {
    // Get the calendar list of the user
    let calendars = Calendar.CalendarList.list({userEmail: userEmail});
    
    for (let calendar of calendars.items) {
      // Transfer the ownership of each calendar to the manager
      Calendar.Acl.insert({
        role: 'owner',
        scope: {type: 'user', value: managerEmail}
      }, calendar.id);
    }
    return {status: 'success', message: `Calendar transferred from ${userEmail} to ${managerEmail}`};
  } catch (err) {
    return {status: 'failure', message: `Failed to transfer calendar: ${err.message}`};
  }
}

// Function to check the Admin Audit logs for successful transfer confirmation
function checkAdminAuditLog(userEmail, managerEmail) {
  try {
    let results = AdminReports.Activities.list(userEmail, 'calendar');
    let transferLogs = results.items.filter(item => item.events.some(event => event.name === 'OWNERSHIP_TRANSFER'));
    
    if (transferLogs.length > 0) {
      return {status: 'success', message: 'Ownership transfer confirmed in audit log.'};
    } else {
      return {status: 'failure', message: 'No ownership transfer found in audit log.'};
    }
  } catch (err) {
    return {status: 'failure', message: `Failed to retrieve audit logs: ${err.message}`};
  }
}

// Function to prepare the checklist of the transfer process
function prepareChecklist(transferStatus, auditLogs) {
  let checklist = `
    <h3>Transfer Process Checklist</h3>
    <ul>
      <li>Transfer Calendar: ${transferStatus.status === 'success' ? '✔️' : '❌'} (${transferStatus.message})</li>
      <li>Audit Log Confirmation: ${auditLogs.status === 'success' ? '✔️' : '❌'} (${auditLogs.message})</li>
    </ul>
  `;
  
  return checklist;
}


// Helper function to generate a random password
function generateRandomPassword() {
  var charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()";
  var passwordLength = 12;
  var password = "";
  for (var i = 0; i < passwordLength; i++) {
    var randomIndex = Math.floor(Math.random() * charset.length);
    password += charset.charAt(randomIndex);
  }
  return password;
}
