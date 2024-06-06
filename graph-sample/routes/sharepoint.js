// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const router = require('express-promise-router').default();
const graph = require('../graph.js');

const { body, validationResult } = require('express-validator');
const validator = require('validator');

/* GET /calendar */
// <GetRouteSnippet>
router.get('/',
  async function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/');
    } else {
      const params = {
        active: { sharepoint: true }
      };
 
      try {
        // Get the events
        const items = await graph.getSharepointItems(
          req.app.locals.msalClient,
          req.session.userId
        );
        // Assign the events to the view parameters
        params.items = items.value;

      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not fetch sharepoint root',
          debug: JSON.stringify(err, Object.getOwnPropertyNames(err))
        });
      }


      res.render('sharepoint', params);
    }
  }
);
// </GetRouteSnippet>

// <GetEventFormSnippet>
/* GET /calendar/new */
router.get('/new',
  function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/');
    } else {
      res.locals.newItem = {};
      res.render('newitem');
    }
  }
);

router.get('/update/:id',
  function(req, res) {
    if (!req.session.userId) {
      // Redirect unauthenticated requests to home page
      res.redirect('/');
    } else {
      const itemId = req.params.id;
      res.locals.updateItem = {};
      res.render('updateitem', {id: itemId});
    }
  });
// </GetEventFormSnippet>
// <PostEventFormSnippet>
/* POST /calendar/new */
router.post('/new', [
  body('title').escape(),
], async function(req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect('/');
  } else {
    // Build an object from the form values
    const formData = {
      title: req.body['title'],
    };

    // Check if there are any errors with the form values
    const formErrors = validationResult(req);
    if (!formErrors.isEmpty()) {

      let invalidFields = '';
      formErrors.array().forEach(error => {
        if (error.type == 'field') {
          invalidFields += `${error.path.slice(3, error.path.length)},`;
        }
      });

      // Preserve the user's input when re-rendering the form

      return res.render('newitem', {
        newItem: formData,
        error: [{ message: `Invalid input in the following fields: ${invalidFields}` }]
      });
    }


    // Create the event
    try {
      await graph.createItem(
        req.app.locals.msalClient,
        req.session.userId,
        formData);
    } catch (error) {
      req.flash('error_msg', {
        message: 'Could not create event',
        debug: JSON.stringify(error, Object.getOwnPropertyNames(error))
      });
    }

    // Redirect back to the calendar view
    return res.redirect('/sharepoint');
  }
}
);



router.delete('/:id', async function(req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect('/');
  } else {
    // Create the event
    const id = req.params.id;

    try {
      await graph.deleteItem(
        req.app.locals.msalClient,
        req.session.userId,
        id);
      
    } catch (error) {
      req.flash('error_msg', {
        message: 'Could not delete item',
        debug: JSON.stringify(error, Object.getOwnPropertyNames(error))
      });
    }

    return res.redirect('/sharepoint');

  }
});

router.put('/:id',[
  body('title').escape(),
], async function(req, res) {
  if (!req.session.userId) {
    // Redirect unauthenticated requests to home page
    res.redirect('/');
  } else {
    // Create the event
    const id = req.params.id;

    const formData = {
      title: req.body['title'],
    };

    try {
      await graph.updateItem(
        req.app.locals.msalClient,
        req.session.userId,
        formData,
        id);

    } catch (error) {
      req.flash('error_msg', {
        message: 'Could not delete item',
        debug: JSON.stringify(error, Object.getOwnPropertyNames(error))
      });
    }

    return res.redirect('/sharepoint');
  }
});
// </PostEventFormSnippet>
module.exports = router;
