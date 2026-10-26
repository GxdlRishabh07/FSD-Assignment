import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def add_hyperlink(paragraph, url, text, color="2B6CB0", underline=True):
    """
    A function that places a hyperlink within a paragraph object.
    """
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')

    if color:
        c = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
        rPr.append(c)
    if underline:
        u = parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>')
        rPr.append(u)

    new_run.append(rPr)
    new_run_text = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(new_run_text)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def format_code_block(doc, code_text):
    """Creates a beautifully styled code box table in Word."""
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.columns[0].width = Inches(6.5)
    
    cell = table.cell(0, 0)
    # Shading (light grey background)
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F7FAFC"/>')
    cell._tc.get_or_add_tcPr().append(shading)
    
    # Border styling (subtle grey border)
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E0"/>'
        f'<w:left w:val="single" w:sz="18" w:space="0" w:color="3182CE"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="CBD5E0"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="CBD5E0"/>'
        f'</w:tcBorders>'
    )
    cell._tc.get_or_add_tcPr().append(borders)
    
    # Cell margins / padding
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="160" w:type="dxa"/>'
        f'<w:left w:w="240" w:type="dxa"/>'
        f'<w:bottom w:w="160" w:type="dxa"/>'
        f'<w:right w:w="240" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(code_text.strip())
    run.font.name = 'Consolas'
    run.font.size = Pt(9.5)
    run.font.color.rgb = RGBColor(45, 55, 72)

def generate_report():
    doc = docx.Document('FSD LAB Report_backup.docx')
    
    # 1. Update Title and Subtitle
    doc.paragraphs[0].text = "FSD LAB REPORT"
    doc.paragraphs[1].text = "Full Stack Web Development - Practical Lab Assignments (1 to 6)"
    
    # Update Student Details
    p3 = doc.paragraphs[3]
    p3.text = ""
    r3 = p3.add_run("Rishabh Patil | PRN: 1272250228")
    r3.bold = True
    r3.font.size = Pt(16)
    r3.font.color.rgb = RGBColor(0x78, 0x3F, 0x04)
    
    p4 = doc.paragraphs[4]
    p4.text = ""
    r4 = p4.add_run("Date: 26.10.2026")
    r4.font.color.rgb = RGBColor(0x78, 0x3F, 0x04)
    
    p5 = doc.paragraphs[5]
    p5.text = ""
    r5 = p5.add_run("FULL STACK WEB DEVELOPMENT LAB")
    r5.bold = True
    r5.font.color.rgb = RGBColor(0x78, 0x3F, 0x04)
    
    # Remove existing dummy paragraphs (from index 10 to end)
    while len(doc.paragraphs) > 10:
        p_elem = doc.paragraphs[-1]._p
        p_elem.getparent().remove(p_elem)
        
    repo_url = "https://github.com/GxdlRishabh07/FSD-Assignment"
    live_url = "https://gxdlrishabh07.github.io/FSD-Assignment/"

    # Data for all 6 assignments
    assignments = [
        {
            "num": 1,
            "title": "Assignment 1: JavaScript Calculator using Functions",
            "objective": "Create a calculator using JavaScript functions to perform basic arithmetic operations (addition, subtraction, multiplication, division). Use if-else or switch statements.",
            "theory": (
                "In this assignment, a modular approach is implemented to build a functional calculator in JavaScript. "
                "Instead of monolithic scripting, dedicated functions are created for each arithmetic operation: add(), subtract(), multiply(), and divide(). "
                "A central controller function performCalculation() extracts the input numbers from the HTML DOM, parses strings to numerical float values using parseFloat(), "
                "and routes execution through a switch statement based on the selected operator (+, -, *, /). "
                "Input validation verifies that fields are non-empty and handles division by zero to ensure application robustness."
            ),
            "code": (
"""// --- Dedicated Arithmetic Functions ---
function add(num1, num2) {
    return num1 + num2;
}

function subtract(num1, num2) {
    return num1 - num2;
}

function multiply(num1, num2) {
    return num1 * num2;
}

function divide(num1, num2) {
    if (num2 === 0) {
        return "Cannot divide by zero"; // Handling zero division
    }
    return num1 / num2;
}

// --- Main Calculator Controller Function ---
function performCalculation() {
    var num1Input = document.getElementById('calc-num1').value;
    var num2Input = document.getElementById('calc-num2').value;
    var operator = document.getElementById('calc-operator').value;
    var resultDisplay = document.getElementById('calc-result');

    // Validation: Check for empty input fields
    if (num1Input === "" || num2Input === "") {
        resultDisplay.innerText = "Please enter both numbers!";
        resultDisplay.style.color = "#e53e3e";
        return;
    }

    var n1 = parseFloat(num1Input);
    var n2 = parseFloat(num2Input);
    var calculatedResult = 0;

    // Switch statement to direct arithmetic operation
    switch (operator) {
        case "+":
            calculatedResult = add(n1, n2);
            break;
        case "-":
            calculatedResult = subtract(n1, n2);
            break;
        case "*":
            calculatedResult = multiply(n1, n2);
            break;
        case "/":
            calculatedResult = divide(n1, n2);
            break;
        default:
            calculatedResult = "Invalid Operator";
    }

    resultDisplay.style.color = "#2b6cb0";
    resultDisplay.innerText = "Result: " + calculatedResult;
}"""
            )
        },
        {
            "num": 2,
            "title": "Assignment 2: DOM-based To-Do List",
            "objective": "Build an interactive to-do list using HTML, CSS, and JavaScript. Use DOM methods to: (a) Add tasks, (b) Mark them as completed, and (c) Delete tasks dynamically.",
            "theory": (
                "The Document Object Model (DOM) is an application programming interface for HTML and XML documents. "
                "This assignment demonstrates dynamic DOM node creation, tree modification, and event handling without page refreshes. "
                "When the user inputs a task and clicks 'Add Task', document.createElement('li') initializes a new list item node. "
                "Child nodes—including a text span and action buttons—are attached using element.appendChild(). "
                "Clicking 'Done' triggers li.classList.toggle('completed') to toggle visual strike-through styling via CSS, "
                "and clicking 'Delete' executes li.remove() to cleanly detach the element from the DOM tree."
            ),
            "code": (
"""function addTask() {
    var inputField = document.getElementById('todo-input');
    var taskText = inputField.value.trim();

    if (taskText === "") {
        alert("Please enter a task!");
        return;
    }

    // Step 1: Create <li> node dynamically
    var li = document.createElement('li');
    li.className = 'todo-item';

    // Step 2: Create task text node
    var span = document.createElement('span');
    span.className = 'todo-text';
    span.innerText = taskText;

    // Step 3: Create actions container
    var actionsDiv = document.createElement('div');
    actionsDiv.className = 'todo-actions';

    // Step 4: Create 'Done/Toggle' action button
    var doneBtn = document.createElement('button');
    doneBtn.className = 'btn-action btn-toggle';
    doneBtn.innerText = 'Done';
    doneBtn.onclick = function() {
        li.classList.toggle('completed'); // Toggle completion class
        doneBtn.innerText = li.classList.contains('completed') ? 'Undo' : 'Done';
    };

    // Step 5: Create 'Delete' action button
    var deleteBtn = document.createElement('button');
    deleteBtn.className = 'btn-action btn-delete';
    deleteBtn.innerText = 'Delete';
    deleteBtn.onclick = function() {
        li.remove(); // Dynamically remove node from DOM
    };

    // Step 6: Assemble element tree and append to unordered list
    actionsDiv.appendChild(doneBtn);
    actionsDiv.appendChild(deleteBtn);
    li.appendChild(span);
    li.appendChild(actionsDiv);
    document.getElementById('todo-list').appendChild(li);

    // Step 7: Clear input field for next entry
    inputField.value = "";
}"""
            )
        },
        {
            "num": 3,
            "title": "Assignment 3: JavaScript Form Validation",
            "objective": "Create a user registration form with fields like name, email, phone, and password. Validate inputs using JavaScript and show appropriate error messages.",
            "theory": (
                "Form validation is essential for data integrity, user experience, and security prior to client-side data submission. "
                "This assignment implements real-time client-side validation using vanilla JavaScript. "
                "The submission event is intercepted using event.preventDefault() to avoid unwanted page reloads. "
                "Validation rules include: Full Name checking for non-emptiness and minimum character length (3+ chars); "
                "Email verification utilizing Regular Expressions (RegEx) matching standard format patterns; "
                "Phone Number validation requiring exactly 10 numerical digits (/^[0-9]{10}$/); "
                "and Password validation requiring a minimum length of 6 characters. "
                "Specific contextual error messages and visual boundary highlights (input-error CSS class) are rendered when constraints fail."
            ),
            "code": (
"""function validateRegistrationForm(event) {
    event.preventDefault(); // Prevent default browser form submission

    var nameInput = document.getElementById('reg-name');
    var emailInput = document.getElementById('reg-email');
    var phoneInput = document.getElementById('reg-phone');
    var passwordInput = document.getElementById('reg-password');

    var errorName = document.getElementById('error-name');
    var errorEmail = document.getElementById('error-email');
    var errorPhone = document.getElementById('error-phone');
    var errorPassword = document.getElementById('error-password');
    var successBox = document.getElementById('reg-success-message');

    // Reset previous errors and highlights
    errorName.innerText = "";
    errorEmail.innerText = "";
    errorPhone.innerText = "";
    errorPassword.innerText = "";
    nameInput.classList.remove('input-error');
    emailInput.classList.remove('input-error');
    phoneInput.classList.remove('input-error');
    passwordInput.classList.remove('input-error');
    successBox.style.display = "none";

    var isValid = true;

    // 1. Full Name Validation
    var nameValue = nameInput.value.trim();
    if (nameValue === "") {
        errorName.innerText = "Full Name is required.";
        nameInput.classList.add('input-error');
        isValid = false;
    } else if (nameValue.length < 3) {
        errorName.innerText = "Name must be at least 3 characters long.";
        nameInput.classList.add('input-error');
        isValid = false;
    }

    // 2. Email Validation with RegEx
    var emailValue = emailInput.value.trim();
    var emailRegex = /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/;
    if (emailValue === "") {
        errorEmail.innerText = "Email address is required.";
        emailInput.classList.add('input-error');
        isValid = false;
    } else if (!emailRegex.test(emailValue)) {
        errorEmail.innerText = "Please enter a valid email address.";
        emailInput.classList.add('input-error');
        isValid = false;
    }

    // 3. Phone Number Validation (10 digits)
    var phoneValue = phoneInput.value.trim();
    var phoneRegex = /^[0-9]{10}$/;
    if (phoneValue === "") {
        errorPhone.innerText = "Phone number is required.";
        phoneInput.classList.add('input-error');
        isValid = false;
    } else if (!phoneRegex.test(phoneValue)) {
        errorPhone.innerText = "Phone number must be exactly 10 digits.";
        phoneInput.classList.add('input-error');
        isValid = false;
    }

    // 4. Password Validation (min 6 characters)
    var passwordValue = passwordInput.value;
    if (passwordValue === "") {
        errorPassword.innerText = "Password is required.";
        passwordInput.classList.add('input-error');
        isValid = false;
    } else if (passwordValue.length < 6) {
        errorPassword.innerText = "Password must be at least 6 characters long.";
        passwordInput.classList.add('input-error');
        isValid = false;
    }

    // Display confirmation on success
    if (isValid) {
        successBox.style.display = "block";
        document.getElementById('registration-form').reset();
    }
}"""
            )
        },
        {
            "num": 4,
            "title": "Assignment 4: React Profile Card using Props",
            "objective": "Set up a React app and create a component to display a user's profile card. Use props to pass name, image URL, and short description.",
            "theory": (
                "React is a component-driven declarative JavaScript library for user interfaces. "
                "In React, components are reusable building blocks that accept arbitrary inputs called 'props' (short for properties) "
                "and return React elements describing what should appear on the screen. "
                "This assignment establishes a reusable ProfileCard component that dynamically receives attributes (name, role, imageUrl, description) "
                "from a parent container (ProfileCardApp). "
                "Props enforce unidirectional data flow from parent to child, promoting maintainable, modular, and testable code."
            ),
            "code": (
"""// --- Reusable ProfileCard Component Receiving Props ---
function ProfileCard(props) {
    return (
        <div className="profile-card">
            {/* Image passed via prop */}
            <img 
                src={props.imageUrl} 
                alt={props.name} 
                className="profile-avatar" 
            />
            {/* Name passed via prop */}
            <h3 className="profile-name">{props.name}</h3>
            {/* Role badge */}
            <div className="profile-role">{props.role}</div>
            {/* Description passed via prop */}
            <p className="profile-desc">{props.description}</p>
        </div>
    );
}

// --- Parent Container Demonstrating Component Reusability ---
function ProfileCardApp() {
    return (
        <div className="cards-grid">
            {/* Card Instance 1 */}
            <ProfileCard 
                name="Rishabh Patil"
                role="Frontend Developer"
                imageUrl="https://images.unsplash.com/photo-1534528741775-53994a69daeb?w=200&auto=format&fit=crop&q=80"
                description="Enthusiastic developer who loves creating clean UI designs, learning modern web technologies and building interactive web apps."
            />

            {/* Card Instance 2 */}
            <ProfileCard 
                name="Aditi Sharma"
                role="React UI Designer"
                imageUrl="https://images.unsplash.com/photo-1494790108377-be9c29b29330?w=200&auto=format&fit=crop&q=80"
                description="Passionate about designing intuitive digital user interfaces, responsive layouts, and accessible component libraries."
            />
        </div>
    );
}

// Render component into DOM container
const root = ReactDOM.createRoot(document.getElementById('react-profile-card-root'));
root.render(<ProfileCardApp />);"""
            )
        },
        {
            "num": 5,
            "title": "Assignment 5: Controlled React Form",
            "objective": "Design a controlled form in React that accepts user inputs and displays the entered data in real-time below the form.",
            "theory": (
                "In React, a controlled component is an input form element whose value is governed by React state rather than internal DOM state. "
                "The React component that renders the form also controls what happens in that form upon subsequent user input. "
                "In this assignment, a formData state object is created using React.useState to manage multiple fields: fullName, email, techStack, and feedback. "
                "When a keystroke occurs, the handleChange event handler updates state using the computed property syntax [name]: value "
                "while preserving other fields with the spread operator (...prev). "
                "Because state dictates the view, the live preview section beneath the form automatically re-renders instantaneously with zero latency."
            ),
            "code": (
"""function ControlledFormApp() {
    // 1. Single state object controlling all input fields
    const [formData, setFormData] = React.useState({
        fullName: '',
        email: '',
        techStack: 'React',
        feedback: ''
    });

    // 2. Generic change handler to synchronize state with inputs
    const handleChange = (e) => {
        const { name, value } = e.target;
        setFormData(prev => ({
            ...prev,
            [name]: value
        }));
    };

    return (
        <div className="controlled-form-container">
            <form onSubmit={(e) => e.preventDefault()}>
                {/* Full Name Controlled Input */}
                <div className="form-group">
                    <label>Full Name:</label>
                    <input 
                        type="text" 
                        name="fullName" 
                        value={formData.fullName} 
                        onChange={handleChange} 
                        placeholder="Enter full name"
                    />
                </div>

                {/* Email Controlled Input */}
                <div className="form-group">
                    <label>Email Address:</label>
                    <input 
                        type="email" 
                        name="email" 
                        value={formData.email} 
                        onChange={handleChange} 
                        placeholder="Enter email"
                    />
                </div>

                {/* Controlled Select Dropdown */}
                <div className="form-group">
                    <label>Favorite Tech Stack:</label>
                    <select name="techStack" value={formData.techStack} onChange={handleChange}>
                        <option value="React">React.js</option>
                        <option value="Node.js">Node.js</option>
                        <option value="JavaScript">Pure JavaScript</option>
                        <option value="Python">Python</option>
                    </select>
                </div>

                {/* Controlled Textarea */}
                <div className="form-group">
                    <label>Bio / Notes:</label>
                    <textarea 
                        name="feedback" 
                        rows="3" 
                        value={formData.feedback} 
                        onChange={handleChange}
                    ></textarea>
                </div>
            </form>

            {/* Real-time live display of state */}
            <div className="realtime-preview-box">
                <h4>⚡ Real-Time Live Preview (Controlled by State):</h4>
                <p><strong>Full Name:</strong> {formData.fullName || "(Waiting for input...)"}</p>
                <p><strong>Email:</strong> {formData.email || "(Waiting for input...)"}</p>
                <p><strong>Selected Tech:</strong> {formData.techStack}</p>
                <p><strong>Bio:</strong> {formData.feedback || "(Waiting for input...)"}</p>
            </div>
        </div>
    );
}

const root = ReactDOM.createRoot(document.getElementById('react-controlled-form-root'));
root.render(<ControlledFormApp />);"""
            )
        },
        {
            "num": 6,
            "title": "Assignment 6: React Counter App using useState",
            "objective": "Create a counter application in React with buttons to increment, decrement, and reset the value. Use React Hooks (useState).",
            "theory": (
                "Hooks are functions that let developers 'hook into' React state and lifecycle features from functional components. "
                "The useState hook is the primary state management primitive in React. It accepts an initial state value and returns an array "
                "with two elements: the current state value, and a state updater function. "
                "In this Counter application, const [count, setCount] = React.useState(0) establishes numerical state. "
                "Three handler functions (handleIncrement, handleDecrement, and handleReset) invoke setCount, which schedules a component re-render. "
                "Dynamic CSS classes are applied conditionally to present color-coded visual cues (green for positive, red for negative, blue for zero)."
            ),
            "code": (
"""function CounterApp() {
    // 1. State variable: count and updater function: setCount
    const [count, setCount] = React.useState(0);

    // 2. Increment logic (+1)
    const handleIncrement = () => {
        setCount(prevCount => prevCount + 1);
    };

    // 3. Decrement logic (-1)
    const handleDecrement = () => {
        setCount(prevCount => prevCount - 1);
    };

    // 4. Reset logic (0)
    const handleReset = () => {
        setCount(0);
    };

    // 5. Dynamic color class calculation
    let colorClass = "counter-zero";
    if (count > 0) colorClass = "counter-positive";
    else if (count < 0) colorClass = "counter-negative";

    return (
        <div className="counter-container">
            <p>Current Value (Managed by <code>useState</code>):</p>

            {/* Numerical Display */}
            <div className={`counter-display ${colorClass}`}>
                {count}
            </div>

            {/* Counter Control Buttons */}
            <div className="counter-buttons">
                <button className="counter-btn btn-decrement" onClick={handleDecrement}>
                    - Decrement
                </button>
                <button className="counter-btn btn-reset" onClick={handleReset}>
                    ↺ Reset
                </button>
                <button className="counter-btn btn-increment" onClick={handleIncrement}>
                    + Increment
                </button>
            </div>
        </div>
    );
}

const root = ReactDOM.createRoot(document.getElementById('react-counter-root'));
root.render(<CounterApp />);"""
            )
        }
    ]

    for item in assignments:
        # Page break before each assignment
        doc.add_page_break()
        
        # Heading 1: Assignment title
        h1 = doc.add_paragraph(item["title"], style="Heading 1")
        h1.paragraph_format.space_before = Pt(12)
        h1.paragraph_format.space_after = Pt(6)
        
        # 1. Problem Statement / Objective
        h2_obj = doc.add_paragraph("1. Problem Statement & Objective", style="Heading 2")
        h2_obj.paragraph_format.space_before = Pt(10)
        h2_obj.paragraph_format.space_after = Pt(4)
        
        p_obj = doc.add_paragraph(item["objective"], style="normal")
        p_obj.paragraph_format.space_after = Pt(8)
        p_obj.paragraph_format.line_spacing = 1.15
        
        # 2. Key Concepts & Explanation
        h2_theory = doc.add_paragraph("2. Key Concepts & Technical Explanation", style="Heading 2")
        h2_theory.paragraph_format.space_before = Pt(10)
        h2_theory.paragraph_format.space_after = Pt(4)
        
        p_theory = doc.add_paragraph(item["theory"], style="normal")
        p_theory.paragraph_format.space_after = Pt(8)
        p_theory.paragraph_format.line_spacing = 1.15
        
        # 3. Source Code
        h2_code = doc.add_paragraph("3. Implementation Source Code", style="Heading 2")
        h2_code.paragraph_format.space_before = Pt(10)
        h2_code.paragraph_format.space_after = Pt(4)
        
        format_code_block(doc, item["code"])
        
        # 4. GitHub Links
        h2_links = doc.add_paragraph("4. GitHub & Live Verification Links", style="Heading 2")
        h2_links.paragraph_format.space_before = Pt(12)
        h2_links.paragraph_format.space_after = Pt(4)
        
        p_repo = doc.add_paragraph(style="normal")
        p_repo.paragraph_format.space_after = Pt(3)
        r_r_label = p_repo.add_run("• Repository Link: ")
        r_r_label.bold = True
        add_hyperlink(p_repo, repo_url, repo_url)
        
        p_live = doc.add_paragraph(style="normal")
        p_live.paragraph_format.space_after = Pt(12)
        r_l_label = p_live.add_run("• GitHub Pages Live Preview: ")
        r_l_label.bold = True
        add_hyperlink(p_live, live_url, live_url)

    doc.save("FSD LAB Report.docx")
    print("FSD LAB Report.docx generated successfully!")

if __name__ == "__main__":
    generate_report()
