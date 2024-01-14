import fetch from 'node-fetch'
import readline from "readline";
import XLSX from 'xlsx'


/**
 * Change json file into Excel spreadsheet 'reports.xlsx'
 *
 * @param {json} json Json object to be changed into the spreadsheet
 * Example data:
 * [
 * {
 *     'S-number': number,
 *     'Programmingsignoffs': number_of_points,
 *     'Designsignoffs': number_of_points
 * },
 * {
 *     'S-number': number,
 *     'Programmingsignoffs': number_of_points,
 *     'Designsignoffs': number_of_points
 * }
 * ]
 *
 *
 */
function fromJsonToExcel(json) {

        var ws = XLSX.utils.json_to_sheet(json);
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "People");
        XLSX.writeFile(wb,'reports.xlsx');
}


//Add your cookie from horus @todo
const cookie = '';

//Adjust number of points got from each assignemnt @todo
const pointsForAssignment = 1;


// Group names excluded from getting points Example group names: Pairs Minor-13, Pairs NEDAP-1, Pairs Resit-21, Pairs Premaster-2 @todo
const blackList = ['Minor', 'NEDAP', 'Premaster', 'Resit']


// Deadline dates for each week @todo
const assignmentsDates = {
        week1: (new Date('2022-11-21')).toISOString().split('T')[0],
        week2: (new Date('2022-11-28')).toISOString().split('T')[0],
        week3: (new Date('2022-12-05')).toISOString().split('T')[0],
        week4: (new Date('2022-12-12')).toISOString().split('T')[0],
        week5: (new Date('2022-12-19')).toISOString().split('T')[0],
        week6: (new Date('2023-01-09')).toISOString().split('T')[0],
        week7: (new Date('2023-01-16')).toISOString().split('T')[0]
}


/**
 * Get id and assignment name of each existing assignment from either Programming or Design
 *
 * @param {string} auth Authorization token
 * @param {string} part Either 'Programming' or 'Design'
 * @returns {Promise<Array>} Returns list of json object, consisting of id and assignment name
 *
 * Example returned data:
 * [ {id: 1221, assignmentName: 'P-1.6}, {id: 1222, assignemntName: 'P-1.8'}]
 **/
async function getAssignments(auth, part) {


    let number;
    //Id of programming page on horus, may be different next year @todo
    if(part === 'Programming') number = '170';
    //Id of design page on horus, may be different next year @todo

    else if(part === 'Design') number = '171';
    else throw new Error('Part is not recognized')


    return await fetch('https://horus.apps.utwente.nl/api/assignmentSets/' + number, {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': auth,
            'Cookie': cookie,
        }
    }).then((response) => {
        if (response.ok) return response.json();
        else if (response.status === 401) throw new Error('Unauthorized');

    }).then((data) => {
        let assignment = [];

        Object.entries(data.assignments).forEach((entry) => {
            let ass = {
                id: entry[1].id,
                assignmentName: entry[1].name
            }
            assignment.push(ass);
        })
        return assignment;
    }).catch(error => {
        if (error instanceof Error && error.message === 'Unauthorized' ) throw new Error('Unauthorized');
        else console.log('Error from getAssignments: ' + error)
    })
}

/**
 *
 * Get id, student number, externalID, group name, group set id and group set external ID for each student (except the black list)
 *
 *
 * @param {string} auth Authorization token
 * @returns {Promise<Array>} Return array of json objects of information about all of the students
 *
 * Example data:
 *
 * [
 *  {
 *      id: 00001,
 *      studentID: 's3008432',
 *      externalId: '132432',
 *      groupName: 'Pairs Green-23',
 *      groupSetID: 23453,
 *      groupSetExternalID: '12437'
 *  },
 *  {
 *      id: 00001,
 *      studentID: 's300843',
 *     externalId: '234124',
 *      groupName: 'Pairs Green-43,
 *      groupSetID: 5434,
 *      groupSetExternalID: '12327'
 *  }
 *
 *
 *
 * ]
 */
async function getStudents(auth) {
    return await fetch('https://horus.apps.utwente.nl/api/groupSets/2514/groups', {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': auth,
            'Cookie': cookie
        }
    }).then((response) => {
        if (response.ok) return response.json()
        else if(response.status === 401) throw new Error('Unauthorized')
    }).then((data) => {

        let students = []

        Object.entries(data).forEach((entry) => {
            Object.entries(entry[1].participants).forEach((person) => {
                //Add only if the students is not in the black list
                if(!blackList.includes(entry[1].name.split('Pairs ')[1].split('-')[0]))
                {
                    let student = {
                        id: person[1].id,
                        studentID: person[1].person.loginId,
                        externalId: entry[1].externalId,
                        groupName: entry[1].name,
                        groupSetID: entry[1].groupSet.id,
                        groupSetExternalID: entry[1].groupSet.externalId
                    }
                    students.push(student)
                }
            })
        })
        return students;
    }).catch(error => {
        if (error instanceof Error && error.message === 'Unauthorized') throw new Error('Unauthorized');
        else console.log('Error from getStudents: ' + error)
    })
}

/**
 *
 * Get date for assignment with id 'participantId' and for student with number 'participantId'
 *
 * @param {number} participantId id from student object
 * @param {number} assignmentId id from assignment object
 * @param {string} auth Authorization token
 * @returns {Promise<string>} Return date, when the assignment was signed-off by particular person
 */
async function getDate(participantId, assignmentId, auth) {
    return fetch('https://horus.apps.utwente.nl/api/signoff/history?participantId='+participantId+'&assignmentId='+assignmentId, {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': auth,
            'Cookie': cookie
        }
    }).then((response) =>
    {
        if(response.ok) return response.json()
        else if(response.status === 401) throw new Error('Unauthorized')
    }).then((data) => {return new Date(data[0].signedAt).toISOString().split('T')[0]
    }).catch(error => {
        if (error instanceof TypeError && error.message === 'Cannot read properties of undefined (reading \'signedAt\')') {}
        else if (error instanceof Error && error.message === 'Unauthorized') throw new Error('Unauthorized');
        else console.log('Error from getDate: ' + error)
    })
}

//Array, where json object will be added to be later changed into the Excel sheet
let toExcel = [];

/**
 *
 *  Add points to the student, if the given student is not in the toExcel array, or toExcel array is empty, then add the student to it,
 *  if the student is in the array add the pointsForAssignment number to the either 'Programming' or 'Design' points
 *
 * @param {json} student Student object
 * @param {string} part Either 'Programming' or 'Design'
 *
 */
function addHousePointsToPerson(student, part) {
    try {
        let student_number = student.studentID;


        // Find student with particular student number in the toExcel array, if there is no such student return undefined
        let thatStudent = toExcel.find(item => item['S-number'] === student_number);

        if(toExcel.length !== 0 && thatStudent !== undefined) {
            if(part === 'Programming') thatStudent.Programmingsignoffs += pointsForAssignment;
            else if(part === 'Design') thatStudent.Designsignoffs += pointsForAssignment;
            else throw new Error('Part is not recognized')
        } else {
            let json = {
                'S-number': student_number,
                'Programmingsignoffs': 0,
                'Designsignoffs': 0
            }
            if(part === 'Programming') json.Programmingsignoffs += pointsForAssignment;
            else if(part === 'Design') json.Designsignoffs += pointsForAssignment;
            else throw new Error('Part is not recognized')

            toExcel.push(json);
        }
    }
    catch (error) {
        console.log('Error from addHousePointsToPerson: ' + error.message);
        console.log(toExcel)
        console.log(student);
        console.log(toExcel.find(item => item['S-number'] === student_number))

    }
}

/**
 *
 * Gives decision if the given student should get points for given assignment of given part.
 *
 *
 * @param {json} assignment Assignment object
 * @param {json} student Student object
 * @param {string} auth Authorization token
 * @param {string} part Either 'Programming' or 'Design'
 * @returns {Promise<boolean>} Returns true if the given assignment for the given student for given part was signed off before the deadline date given in assignmentsDates
 */
async function decision(assignment, student, auth, part) {

    let assignmentID;

    if(part === 'Programming') assignmentID = assignment.assignmentName.split('P-')[1].split('.')[0];
    else if(part === 'Design') assignmentID = assignment.assignmentName.split('D-')[1].split('.')[0];
    else throw new Error('part is not recognized')

    let finalDate;

    switch(assignmentID) {
        case '1':
            finalDate = assignmentsDates.week1;
            break;
        case '2':
            finalDate = assignmentsDates.week2;
            break;
        case '3':
            finalDate =assignmentsDates.week3;
            break;
        case '4':
            finalDate = assignmentsDates.week4;
            break;
        case '5':
            finalDate = assignmentsDates.week5;
            break;
        case '6':
            finalDate = assignmentsDates.week6;
            break;
        case '7':
            finalDate = assignmentsDates.week7;
            break;
    }
        return await getDate(student.id, assignment.id, auth).then((date) => { return (finalDate >= date);
        }).catch((error) => {
            if (error instanceof TypeError && error.message === 'Cannot read properties of undefined (reading \'signedAt\')') {}
            else if (error instanceof Error && error.message === 'Unauthorized')  throw new Error('Unauthorized');
            else console.log('Error from decision: ' + error)
        });
}

/**
 *
 * Remove student with given student number (sid) from the toExcel array
 *
 * @param {string} sid Student number of student that needes to be removed
 * @returns {boolean}
 */
const removeById = (sid) => {
    try{
        const requiredIndex = toExcel.findIndex(el => {return el['S-number'] === sid;});

        if(requiredIndex === -1) throw new Error('Student: ' + sid + ' hasn\'t been removed');
        else toExcel.splice(requiredIndex, 1);

    }catch (error) { console.log('Error from removeById: ' + error.message )}
};

/**
 *
 * Main function which merges all the other functions together and is crucial for developing the code without implementation of refresh token
 *
 * In the begining the function asks you for your auth token, which you need to get from logging into horus and collecting content of Authorization request header from any request sent (preferably it should be 170 request)
 * After the timout, that will occur to refresh the token, application will stop and ask you agan for the new content of Authorization request header and will save all the progress.
 * User then needs to refresh the horus page and collect the content of Authorization request header once again.
 *
 * @param {number} students_start Number of students that have the points already calculated
 * @param {json} last_student The student that was being processed, when the refresh occurred to check if app does not leave anyone. It is necessary as most probably function will break in the middle of
 * calculation points for a student. Thus the varialbe needs to be deleted and calculated once again
 *
 */
async function main(students_start, last_student) {

    // Student start variable ,which will count how many students have been processed after starting the function
    let sstart = 0;

    // Reading auth token from the terminal
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    let auth = await new Promise(resolve  => {
        rl.question("To continue insert auth token: \n", (input) => {
            resolve(input);
        });
    });
    rl.close()



    let students;
    let assignments_programming;
    let assignments_design;
    let current_student
    try {

        // Getting array of json objects of all the students, given the auth token
        students = await getStudents(auth);

        // Getting array of json objects for all programming assignments, given the auth token
        assignments_programming = await getAssignments(auth, 'Programming');

        // Getting array of json objects for all design assignments, given the auth token
        assignments_design = await getAssignments(auth, 'Design');

        // Getting only students that haven't been processed yet
        let studentSet = students.slice(students_start);

        // If the last student is not undefined, then show if the last student from previous iteration of the function has been deleted
        if(last_student !== undefined ) console.log('Last student deleted: ' + !(toExcel.some(item => item['S-number'] === last_student.studentID)));


        // Showing the progress
        let progress = (students_start) / students.length;

        // Loop through each student
        for (let student of Object.values(studentSet)) {
                progress += 1 / students.length;
                console.log(progress * 100 + '%')
                current_student = student;

                // Loop through each programming assignment
                for (let assignment of Object.values(assignments_programming)) {

                    // Checking decision for programming assignment and adding the points
                    if (await decision(assignment, student, auth, 'Programming')) {
                        addHousePointsToPerson(student, 'Programming')
                    }
                }
                // Loop through each design assignment
                for (let assignment of Object.values(assignments_design)) {
                    // Checking decision for design assignment and adding the points=
                    if (await decision(assignment, student, auth, 'Design')) {
                        addHousePointsToPerson(student, 'Design')
                    }
                }
                sstart++;

        }
        // Transfer toExcel array to Excel spreadsheets after finishing the calculations
        fromJsonToExcel(toExcel);

        console.log('All done! \n\nYou can go to \'reports.xlsx\' to get the students and their points')

    }
    catch(error) {

        // If Unauthorized error occures
        if(error instanceof Error && error.message === 'Unauthorized'){
            console.log('Unauthorized from main')

            // If current_student is not undefined, needs to be removed from the toExcel list
            if(current_student !== undefined) removeById(current_student.studentID);

            // Run main once again, starting with studnets_start + sstart and current_students
            main(students_start + sstart, current_student)
        }
        else { console.log('Error from main: ' + error)}
    }
}

//Run main function with student_start 0 and last_student undefined
main(0, undefined);

