/* eslint-disable */
import "../../assets/todoist-16.png";
import "../../assets/todoist-32.png";
import "../../assets/todoist-80.png";
import "../../assets/gantt-16.png";
import "../../assets/gantt-32.png";
import "../../assets/gantt-80.png";

let token = localStorage["todoist_token"] || "none",
tasks = [],
view = {

	get(id) {
		return document.getElementById(id)
	},

	update(data) {
		for (let id in data) {
			let el = document.getElementById(id)
			switch (el.nodeName) {
				case 'INPUT':
					el.value = data[id]
					break
				case 'IMG':
					el.src = data[id]
					break
				default:
					el.innerHTML = data[id]
			}
		}
   },

	alert(alertTitle, alertDetails = "") {
		this.update({ alertTitle, alertDetails })
		this.show("alert")
	},

	hide(id) {
		this.get(id).style.display = 'none'
	},
 
	show(id) {
   		this.get("status").style.display = 'none'
		this.get(id).style.display = 'inline'
	},

	wait() {
		view.hide("connect")
		view.hide("settings")
		view.hide("alert")
		view.show("status")
	},

	connect() {
		//if (view.get("token").value.length == 0) return
		token = view.get("token").value
		localStorage.setItem("todoist_token", token)
		view.hide("connect")
		view.sync()
	},

	sync() {

		let sync_url = "https://api.todoist.com/sync/v8/sync",
		headers = {
			'Authorization': 'Bearer ' + token,
			'Content-Type': 'application/json',
		},

		params = {
			sync_token: '*',
			resource_types: ["all"]
		}

		view.wait()
		fetch(sync_url, { 
			method: 'POST',
			headers : headers,
			body: JSON.stringify(params)
		})
	
		.then(res => {
			res.json().then(data => {
				view.update({
					avatar: data.user.avatar_medium,
					user: data.user.full_name,
					mail: data.user.email
				})

				let list = view.get("projects")
				data.projects.forEach(project => {
					list.options.add(new Option(project.name, project.id))
				})

				let projects = view.get("dropdown")
				new fabric['Dropdown'](projects)
				view.show("settings")
			})
		})

		.catch(error => {
			view.alert("Connection failed.", error);
		})
	},

	pushTasks() {
		
		let item = 0,
		headers = {
			'Authorization': 'Bearer ' + token,
			'Content-Type': 'application/json'
		}

		view.wait()
		tasks.forEach( todo => {
			fetch('https://api.todoist.com/rest/v1/tasks', { 
				method: 'POST',
				headers : headers,
				body: JSON.stringify({ content: todo })
			})

			.then(res => {
				item ++
				console.log("status: ", res.status)
				if (item == tasks.length) {
					Office.context.ui.closeContainer()
					window.open("https://todoist.com", "_blank")
				}
			})

			.catch(error => {
				view.alert("Task"+ item +" push failed!", error);
			})

		})

		
	}

}

Office.onReady((info) => {
	if (info.host === Office.HostType.OneNote) {

		let ToggleElements = document.querySelectorAll(".ms-Toggle");
		for (let i = 0; i < ToggleElements.length; i++) {
			new fabric['Toggle'](ToggleElements[i]);
		}

		let CloseElements = document.querySelectorAll(".close");
		for (let i = 0; i < CloseElements.length; i++) {
			CloseElements[i].addEventListener("click", closeTaskPane)
		}

		view.get("login").onclick = view.connect;
		view.get("submit").onclick = view.pushTasks;
		getPageTasks()

	}
})

export async function getPageTasks() {

	view.wait()
	OneNote.run(context => {
		let parser = new DOMParser(),
		page = context.application.getActivePage(),
		outlines = []

		page.load("contents");
		page.contents.load("items");
		
		return context.sync().then(() => {
			console.log("checking outlines...");
			page.contents.items.forEach(item => {
				outlines.push(item)
				item.outline.paragraphs.load("items");
			})

			return context.sync().then(() => {
				let strings = [];
				console.log("checking paragraphs...");
				outlines.forEach( item => {
					item.outline.paragraphs.items.forEach( p => {
						if (p.type == "RichText"){
							let html = p.richText.getHtml();
							strings.push(html)
							p.load("richtext");
						}
						console.log(p.type)
					})
				})

				return context.sync().then(function(){
					strings.forEach( html => {
						let doc = parser.parseFromString(html.value, 'text/html'),
						tag = doc.querySelector("[data-tag=to-do]")
						if (tag != null) tasks.push(tag.innerText)
					})
					
					console.log('tasks found:', tasks.length)
					view.update({ tasks: tasks.length + " task(s)" })
					token = "none"

					if (tasks.length == 0) {
						view.alert("Sorry, no tasks found!", "No to-do tags found on this page!&emsp;")
					} else if (token == "none") {
						view.show("connect")
					} else {
						view.sync()
					}
					
				})

			})

		})

		.catch(error => {
			view.alert("Sync error!", error);
		})

	})
}

export async function closeTaskPane() {
	Office.context.ui.closeContainer();
}